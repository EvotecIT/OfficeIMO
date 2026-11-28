import { standardCode } from "../../pdf/font-support.js";

const replacements: ReadonlyMap<number, string> = new Map([[0x2212, "-"], [0x2010, "-"], [0x2011, "-"], [0x2012, "-"], [0x2009, " "], [0x200a, " "],
  [0x202f, " "], [0x2007, " "], [0x2002, " "], [0x2003, " "], [0x200b, ""], [0x200e, ""], [0x200f, ""], [0x2060, ""], [0xfeff, ""]]);

/** Keep text within the standard PDF fonts' WinAnsi repertoire: typographic minus and spaces get their ASCII
 * equivalents, invisible marks are dropped and any other character becomes `?`, reported through `substituted`. */
export function winAnsiText(text: string, substituted: () => void): string {
  let result = "";
  for (const character of text) {
    const scalar = character.codePointAt(0)!;
    if (standardCode(scalar) !== undefined) { result += character; continue; }
    if (scalar === 10 || scalar === 13 || scalar === 9) { result += character; continue; }
    const replacement = replacements.get(scalar);
    if (replacement === undefined) substituted();
    result += replacement ?? "?";
  }
  return result;
}
