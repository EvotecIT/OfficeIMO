import type { CellPresentation } from "../../core/index.js";

/** Stable document palette for CanopyX's standard tones. Browser CSS variables and resolved cell.format are never interpreted.
 * Each foreground keeps at least a 4.5:1 contrast ratio on its background. */
export const canopyTonePalette: Readonly<Record<string, { readonly color: string; readonly background?: string; readonly foreground: string }>> = Object.freeze({
  info: Object.freeze({ color: "4472C4", background: "DDEBF7", foreground: "1F4E78" }),
  success: Object.freeze({ color: "70AD47", background: "E2F0D9", foreground: "375623" }),
  warning: Object.freeze({ color: "FFC000", background: "FFF2CC", foreground: "7F5000" }),
  danger: Object.freeze({ color: "C00000", background: "FFC7CE", foreground: "9C0006" }),
  // Muted tones dim the text rather than painting the cell.
  muted: Object.freeze({ color: "A6A6A6", foreground: "595959" })
});

/** Default cell presentation for a standard tone; `neutral` and unknown tones have none. */
export const canopyDefaultTones: ReadonlyMap<string, Readonly<CellPresentation>> = new Map(Object.entries(canopyTonePalette).map(([name, tone]) =>
  [name, Object.freeze({ ...(tone.background === undefined ? {} : { background: tone.background }), color: tone.foreground })] as const));
