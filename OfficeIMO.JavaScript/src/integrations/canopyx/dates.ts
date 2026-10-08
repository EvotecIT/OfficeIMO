import type { CanopySourceOptions } from "./types.js";

/** @internal Typed conversion follows declared datetime intent, never localized display text. */
export function datetime(value: string, mode: CanopySourceOptions["datetime"]): Date | string {
  if (mode === "text") return value;
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,7}))?(Z|[+-]\d{2}:\d{2})$/.exec(value);
  if (!match) throw new TypeError("Canopy datetime values require an ISO timestamp with an explicit time zone.");
  const date = new Date(value);
  const offset = match[8] === "Z" ? 0 : (match[8]![0] === "+" ? 1 : -1) * (Number(match[8]!.slice(1, 3)) * 60 + Number(match[8]!.slice(4)));
  const wall = new Date(date.getTime() + offset * 60000);
  if (!Number.isFinite(date.getTime()) || Number(match[8]!.slice(1, 3)) > 23 || Number(match[8]!.slice(4)) > 59 ||
    wall.getUTCFullYear() !== Number(match[1]) || wall.getUTCMonth() + 1 !== Number(match[2]) || wall.getUTCDate() !== Number(match[3]) ||
    wall.getUTCHours() !== Number(match[4]) || wall.getUTCMinutes() !== Number(match[5]) || wall.getUTCSeconds() !== Number(match[6]))
    throw new TypeError("Invalid Canopy datetime value.");
  if (/[1-9]/.test((match[7] ?? "").slice(3))) {
    if (mode === "typed") throw new TypeError("Typed Excel dates cannot preserve sub-millisecond precision; use datetime: preserve or text.");
    return value;
  }
  if (date.getUTCFullYear() < 1900 || date.getUTCFullYear() > 9999) {
    if (mode === "typed") throw new TypeError("Typed Excel dates require years 1900 through 9999; use datetime: preserve or text.");
    return value;
  }
  return date;
}
