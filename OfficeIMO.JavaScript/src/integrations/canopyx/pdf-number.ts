import { ExportCell } from "../../core/index.js";
import type { ExportValue, ColumnSettings } from "../../core/index.js";

/** Fixed numeric formats affect PDF text, not the frozen numeric value used by other writers. */
function numericProfile(format: string | undefined): { minimumFractionDigits: number; maximumFractionDigits: number; percent: boolean; grouping: boolean; prefix: string; suffix: string } | undefined {
  if (format === undefined) return undefined;
  const match = /^((?:"[^"\\]*")*)(#,##)?0(?:\.([0#]+))?(%)?((?:"[^"\\]*")*)$/.exec(format);
  if (!match || (match[3]?.length ?? 0) > 10) return undefined;
  return { minimumFractionDigits: (match[3]?.match(/0/g) ?? []).length, maximumFractionDigits: match[3]?.length ?? 0,
    percent: !!match[4], grouping: !!match[2], prefix: match[1]!.replace(/"/g, ""), suffix: match[5]!.replace(/"/g, "") };
}

/** Preserve the captured decimal convention across a column, including whole percentages without a separator. */
export function canopyPdfNumbers(sample: readonly (readonly ExportValue[])[], columns: readonly ColumnSettings[]): (row: readonly ExportValue[]) => readonly ExportValue[] {
  const formatters = new Map<string, Intl.NumberFormat>();
  const decimal = columns.map((column, index) => {
    let fallback = ".";
    for (const row of sample) {
      const cell = row[index];
      if (!(cell instanceof ExportCell) || typeof cell.value !== "number") continue;
      const profile = numericProfile(cell.presentation?.numberFormat ?? column.format);
      if (!profile?.maximumFractionDigits) continue;
      const match = /([.,])(\d+)[^\d.,]*$/.exec(cell.text ?? "");
      if (match && match[2]!.length <= profile.maximumFractionDigits) {
        const numeric = /([+-]?\d[\d.,]*)[^\d]*$/.exec(cell.text ?? "")?.[1] ?? "";
        const scaled = cell.value * (profile.percent ? 100 : 1);
        if (/^[+-]?\d{1,3}(?:,\d{3})+$/.test(numeric) || /^[+-]?\d{1,3}(?:\.\d{3})+$/.test(numeric)) {
          if (Math.abs(Number(numeric.replace(/[.,]/g, "")) - scaled) < .5) {
            fallback = match[1] === "," ? "." : ","; continue;
          }
        }
        return match[1];
      }
    }
    return fallback;
  });
  return row => row.map((value, index) => {
    if (!(value instanceof ExportCell) || typeof value.value !== "number") return value;
    const profile = numericProfile(value.presentation?.numberFormat ?? columns[index]!.format);
    if (!profile) return value;
    const captured = /^([^0-9+-]*)([+-]?\d[\d., \u00a0\u202f]*?)([^0-9]*)$/.exec(value.text ?? "");
    if (value.text && !captured) return value;
    const key = [decimal[index], profile.minimumFractionDigits, profile.maximumFractionDigits, profile.grouping].join(":");
    let formatter = formatters.get(key);
    if (!formatter) { formatter = new Intl.NumberFormat(decimal[index] === "," ? "de-DE" : "en-US", {
      minimumFractionDigits: profile.minimumFractionDigits, maximumFractionDigits: profile.maximumFractionDigits, useGrouping: profile.grouping
    }); formatters.set(key, formatter); }
    const formatted = formatter.format(value.value * (profile.percent ? 100 : 1));
    const prefix = profile.prefix || captured?.[1] || "";
    let suffix = profile.suffix || captured?.[3] || "";
    if (profile.percent && !profile.suffix) suffix = suffix.replace(/%/g, "");
    return new ExportCell(value.value, { text: prefix + formatted + (profile.percent ? "%" : "") + suffix,
      ...(value.presentation ? { presentation: value.presentation } : {}), ...(value.link ? { link: value.link } : {}) });
  });
}
