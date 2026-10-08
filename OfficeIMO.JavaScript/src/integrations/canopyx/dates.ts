import type { CanopySourceOptions } from "./types.js";

/** @internal Typed conversion follows declared datetime intent, never localized display text. */
export function datetime(value: string, mode: CanopySourceOptions["datetime"], clock: "utc" | "local"): Date | string {
  if (mode === "text") return value;
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?(Z|[+-]\d{2}:\d{2})$/.exec(value);
  if (!match) throw new TypeError("Canopy datetime values require an ISO timestamp with an explicit time zone.");
  const date = new Date(value);
  const offset = match[8] === "Z" ? 0 : (match[8]![0] === "+" ? 1 : -1) * (Number(match[8]!.slice(1, 3)) * 60 + Number(match[8]!.slice(4)));
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  const hour = Number(match[4]), minute = Number(match[5]), second = Number(match[6] ?? "0"), fraction = match[7] ?? "";
  // setUTCFullYear preserves years 0 through 99. Validate the calendar before allowing ISO's next-day 24:00 spelling.
  const wall = new Date(0); wall.setUTCFullYear(year, month - 1, day);
  if (!Number.isFinite(date.getTime()) || Number(match[8]!.slice(1, 3)) > 23 || Number(match[8]!.slice(4)) > 59 ||
    wall.getUTCFullYear() !== year || wall.getUTCMonth() + 1 !== month || wall.getUTCDate() !== day ||
    hour > 24 || minute > 59 || second > 59 || hour === 24 && (minute !== 0 || second !== 0 || /[1-9]/.test(fraction)))
    throw new TypeError("Invalid Canopy datetime value.");
  wall.setUTCHours(hour, minute, second, Number((fraction + "000").slice(0, 3)));
  if (wall.getTime() !== date.getTime() + offset * 60000) throw new TypeError("Invalid Canopy datetime value.");
  if (/[1-9]/.test(fraction.slice(3))) {
    if (mode === "typed") throw new TypeError("Typed Excel dates cannot preserve sub-millisecond precision; use datetime: preserve or text.");
    return value;
  }
  const excelYear = clock === "utc" ? date.getUTCFullYear() : date.getFullYear();
  if (excelYear < 1900 || excelYear > 9999) {
    if (mode === "typed") throw new TypeError("Typed Excel dates require years 1900 through 9999; use datetime: preserve or text.");
    return value;
  }
  return date;
}
