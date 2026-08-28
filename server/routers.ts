import { COOKIE_NAME } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, router } from "./_core/trpc";
import { inspectFile, processOCR, userFacingOCRMessage } from "./ocr";
import { buildWordPlan } from "@shared/word";
import { buildExcelPlan, operationForFile } from "@shared/excel";
import { z } from "zod";

const ocrInput = z.object({
  fileName: z.string().min(1).max(255),
  mimeType: z.string().min(1).max(120),
  bytesBase64: z.string().min(1).max(18_000_000),
});

export const appRouter = router({
  system: systemRouter,
  auth: router({
    me: publicProcedure.query(opts => opts.ctx.user),
    logout: publicProcedure.mutation(({ ctx }) => {
      const cookieOptions = getSessionCookieOptions(ctx.req);
      ctx.res.clearCookie(COOKIE_NAME, { ...cookieOptions, maxAge: -1 });
      return { success: true } as const;
    }),
  }),
  ocr: router({
    process: publicProcedure.input(ocrInput).mutation(async ({ input }) => {
      try {
        return await processOCR(input);
      } catch (error) {
        throw new Error(userFacingOCRMessage(error));
      }
    }),
  }),
  excel: router({
    transform: publicProcedure.input(ocrInput.extend({ operation: z.enum(["pdf", "image", "smart"]).default("smart"), mode: z.enum(["separate", "single", "smart"]).default("smart") })).mutation(async ({ input }) => {
      try { const detectedOperation = operationForFile(input.fileName, input.mimeType); if (input.operation !== "smart" && input.operation !== detectedOperation) throw new Error("OPERATION_FILE_MISMATCH"); const document = await processOCR(input); return { document, plan: buildExcelPlan(document, input.mode) }; }
      catch (error) { throw new Error(userFacingOCRMessage(error)); }
    }),
  }),
  word: router({
    inspect: publicProcedure.input(ocrInput).mutation(({ input }) => {
      try { return inspectFile(input); } catch (error) { throw new Error(userFacingOCRMessage(error)); }
    }),
    transform: publicProcedure.input(ocrInput).mutation(async ({ input }) => {
      try {
        const document = await processOCR(input);
        return { document, plan: buildWordPlan(document) };
      } catch (error) {
        throw new Error(userFacingOCRMessage(error));
      }
    }),
  }),
});

export type AppRouter = typeof appRouter;
