export type RibbonPanel = "" | "review" | "settings";

export function parseRibbonPanel(search: string): RibbonPanel {
  const panel = new URLSearchParams(search).get("panel");
  return panel === "review" || panel === "settings" ? panel : "";
}

export function ribbonEntryUrl(baseUrl: string, panel: Exclude<RibbonPanel, "">): string {
  const url = new URL(baseUrl);
  url.searchParams.set("panel", panel);
  return url.toString();
}
