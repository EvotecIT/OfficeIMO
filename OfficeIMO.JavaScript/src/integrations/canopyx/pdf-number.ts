import { ExportCell } from "../../core/index.js";
import type { ExportValue, ColumnSettings } from "../../core/index.js";

/** Fixed numeric formats affect PDF text, not the frozen numeric value used by other writers. */
function numericProfile(format: string | undefined): { minimumFractionDigits: number; maximumFractionDigits: number; percent: boolean; grouping: boolean; prefix: string; suffix: string } | undefined {
  // No declared precision: whole values (including counts) stay whole.
  if (format === undefined || format === "General") return { minimumFractionDigits: 0, maximumFractionDigits: 20,
    percent: false, grouping: false, prefix: "", suffix: "" };
  const match = /^((?:"[^"]*"|\\")*)(#,##)?0(?:\.([0#]+))?(%)?((?:"[^"]*"|\\")*)$/.exec(format);
  if (!match || (match[3]?.length ?? 0) > 10) return undefined;
  const literal = (text: string) => text.replace(/"([^"]*)"|\\(")/g, (_, quoted: string | undefined, escaped: string | undefined) => quoted ?? escaped!);
  return { minimumFractionDigits: (match[3]?.match(/0/g) ?? []).length, maximumFractionDigits: match[3]?.length ?? 0,
    percent: !!match[4], grouping: !!match[2], prefix: literal(match[1]!), suffix: literal(match[5]!) };
}

/** Separate declared literal units before recognizing the captured numeric token. */
function capturedNumber(text: string, profile: NonNullable<ReturnType<typeof numericProfile>>): { prefix: string; number: string; suffix: string } | undefined {
  let body = text;
  if (profile.prefix && body.startsWith(profile.prefix)) body = body.slice(profile.prefix.length);
  if (profile.suffix && body.endsWith(profile.suffix)) body = body.slice(0, -profile.suffix.length);
  const match = /^([^0-9+-]*)([+-]?\d[\d., \u00a0\u202f]*?)([^0-9]*)$/.exec(body);
  return match ? { prefix: profile.prefix || match[1]!, number: match[2]!, suffix: profile.suffix || match[3]! } : undefined;
}

/** Choose once for the document; percent text is the strongest captured culture signal. */
function documentDecimal(sample: readonly (readonly ExportValue[])[], columns: readonly ColumnSettings[]): string {
  let fallback = ".";
  for (const percent of [true, false]) {
    for (const row of sample) for (let index = 0; index < columns.length; index++) {
      const cell = row[index], column = columns[index]!;
      if (!(cell instanceof ExportCell) || typeof cell.value !== "number") continue;
      const profile = numericProfile(cell.presentation?.numberFormat ?? column.format);
      if (!profile?.maximumFractionDigits) continue;
      const captured = capturedNumber(cell.text ?? "", profile);
      if (!captured) continue;
      // A literal percent label is a culture signal without requesting numeric scaling.
      const percentageSignal = profile.percent || profile.suffix.includes("%") || (cell.text ?? "").includes("%");
      if (percentageSignal !== percent) continue;
      const match = /([.,])(\d+)$/.exec(captured.number.trimEnd());
      if (match && match[2]!.length <= profile.maximumFractionDigits) {
        const numeric = captured.number;
        const scaled = cell.value * (profile.percent ? 100 : 1);
        if (/^[+-]?\d{1,3}(?:,\d{3})+$/.test(numeric) || /^[+-]?\d{1,3}(?:\.\d{3})+$/.test(numeric)) {
          if (Math.abs(Number(numeric.replace(/[.,]/g, "")) - scaled) < .5) {
            fallback = match[1] === "," ? "." : ","; continue;
          }
        }
        return match[1]!;
      }
    }
  }
  return fallback;
}

/** Apply one captured numeric convention to every supported numeric cell, including later streamed rows. */
export function canopyPdfNumbers(sample: readonly (readonly ExportValue[])[], columns: readonly ColumnSettings[]): (row: readonly ExportValue[]) => readonly ExportValue[] {
  const formatters = new Map<string, Intl.NumberFormat>(), decimal = documentDecimal(sample, columns);
  return row => row.map((value, index) => {
    if (!(value instanceof ExportCell) || typeof value.value !== "number") return value;
    const format = value.presentation?.numberFormat ?? columns[index]!.format;
    let profile = numericProfile(format);
    if (!profile || value.text === "") return value;
    const captured = capturedNumber(value.text ?? "", profile);
    if (value.text && !captured) return value;
    // General overrides only plain numeric labels; it must not reinterpret a
    // captured percentage or other decorated presentation as an unscaled value.
    if (format === "General" && captured && (captured.prefix.trim() || captured.suffix.trim())) return value;
    if (format === undefined || format === "General") {
      const digits = Number.isInteger(value.value) ? 0 : /[.,](\d+)$/.exec(captured?.number ?? "")?.[1]?.length;
      // Retain a fractional value's captured precision, without padding an
      // undeclared whole aggregate merely because its display text was padded.
      if (digits !== undefined && digits <= 20) profile = { ...profile, minimumFractionDigits: digits, maximumFractionDigits: digits };
    }
    const key = [profile.minimumFractionDigits, profile.maximumFractionDigits, profile.grouping].join(":");
    let formatter = formatters.get(key);
    if (!formatter) { formatter = new Intl.NumberFormat(decimal === "," ? "de-DE" : "en-US", {
      minimumFractionDigits: profile.minimumFractionDigits, maximumFractionDigits: profile.maximumFractionDigits, useGrouping: profile.grouping
    }); formatters.set(key, formatter); }
    const formatted = formatter.format(value.value * (profile.percent ? 100 : 1));
    const prefix = captured?.prefix || profile.prefix;
    let suffix = captured?.suffix || profile.suffix;
    if (profile.percent && !profile.suffix) suffix = suffix.replace(/%/g, "");
    return new ExportCell(value.value, { text: prefix + formatted + (profile.percent ? "%" : "") + suffix,
      ...(value.presentation ? { presentation: value.presentation } : {}), ...(value.link ? { link: value.link } : {}) });
  });
}
