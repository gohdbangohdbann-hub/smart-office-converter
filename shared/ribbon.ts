export type OfficeRibbonHost = "Word" | "Excel";
export type RibbonPanel = "import" | "review" | "settings";

export type RibbonCommand = {
  host: OfficeRibbonHost;
  panel: RibbonPanel;
  label: string;
};

export const ribbonCommands: RibbonCommand[] = [
  { host: "Word", panel: "import", label: "استيراد وتحويل" },
  { host: "Word", panel: "review", label: "مراجعة" },
  { host: "Word", panel: "settings", label: "الإعدادات" },
  { host: "Excel", panel: "import", label: "استيراد وتحويل" },
  { host: "Excel", panel: "review", label: "مراجعة" },
  { host: "Excel", panel: "settings", label: "الإعدادات" },
];

export function parseRibbonPanel(value: string | null | undefined): RibbonPanel | null {
  if (value === "import" || value === "review" || value === "settings") return value;
  return null;
}

export function ribbonPanelForHost(host: OfficeRibbonHost, panel: RibbonPanel): RibbonCommand {
  return ribbonCommands.find((command) => command.host === host && command.panel === panel)!;
}
