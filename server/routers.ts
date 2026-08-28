import { COOKIE_NAME } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, router } from "./_core/trpc";
import { processOCR, userFacingOCRMessage } from "./ocr";
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
});

export type AppRouter = typeof appRouter;
