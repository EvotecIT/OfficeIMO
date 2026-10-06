import { OfficeIMOError } from "../core/errors.js";
import { cleanXml } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";

/** @internal Overflow is deliberately bounded: ZIP worksheet entries cannot be interleaved. */
export class TextOverflow {
  private entries: { sheet: string; cell: string; text: string; first: number; parts: number }[] = [];
  private characters = 0;
  private rows = 0;
  constructor(readonly sheetName: string, private readonly maximum: number, private readonly policy: InvalidCharacterPolicy) {}
  add(sheet: string, cell: string, value: string): { preview: string; location: string } {
    const text = cleanXml(value, this.policy);
    if (this.characters + text.length > this.maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "maxOverflowCharacters exceeded; use CSV or raise the bounded preservation limit.");
    let parts = 0;
    for (const _ of splitText(text)) { void _; parts++; }
    if (this.rows + parts > 1048575) throw new OfficeIMOError("RESOURCE_LIMIT", "Text overflow exceeds the worksheet row limit.");
    const first = this.rows + 2;
    this.entries.push({ sheet, cell, text, first, parts }); this.rows += parts; this.characters += text.length;
    const marker = "\n[Full text: " + this.sheetName + "!D" + first + "; " + parts + " parts]";
    return { preview: clipText(text, 32767 - marker.length) + marker, location: "'" + this.sheetName.replace(/'/g, "''") + "'!D" + first };
  }
  *values(): Generator<readonly (string | number)[]> {
    for (const entry of this.entries) {
      let part = 0;
      for (const text of splitText(entry.text)) yield [entry.sheet, entry.cell, ++part, text, entry.parts];
    }
  }
  clear(): void { this.entries = []; this.characters = 0; this.rows = 0; }
}
/** @internal Clip at a UTF-16 boundary without splitting a valid supplementary character. */
export function clipText(text: string, maximum: number): string {
  let end = Math.min(maximum, text.length);
  if (end < text.length && /[\ud800-\udbff]/.test(text[end - 1] ?? "")) end--;
  return text.slice(0, end);
}
function* splitText(text: string): Generator<string> {
  let offset = 0;
  while (offset < text.length) { const part = clipText(text.slice(offset, offset + 32768), 32767); offset += part.length; yield part; }
}
