"""Generate OfficeIMO.Core Unicode text-layout tables from the Unicode Character Database.

The tables back the managed Unicode Bidirectional Algorithm (UAX #9) and line breaking (UAX #14):
Bidi_Class, Bidi_Paired_Bracket(_Type), Line_Break, East_Asian_Width (F/W/H only) and
Extended_Pictographic. Values are stored as sorted range starts with one value per range, so a
lookup is a binary search and unlisted code points take the documented @missing defaults.
Canonical_Combining_Class and Modifier_Combining_Mark support font-aware normalization and AMTRA.

Usage (Python 3, no third-party packages):
    python Build/Generate-UnicodeTextData.py <directory containing the UCD 16.0.0 files>
    python Build/Generate-UnicodeTextData.py <directory> --combining-only

Required files: extracted/DerivedBidiClass.txt (or DerivedBidiClass.txt), BidiBrackets.txt,
LineBreak.txt, EastAsianWidth.txt and emoji/emoji-data.txt (or emoji-data.txt).
Combining tables require extracted/DerivedCombiningClass.txt (or DerivedCombiningClass.txt)
and PropList.txt. --combining-only regenerates only that separate partial.
"""
import hashlib
import re
import sys
from pathlib import Path

BIDI_CLASSES = ["L", "R", "AL", "EN", "ES", "ET", "AN", "CS", "NSM", "BN", "B", "S", "WS", "ON",
                "LRE", "LRO", "RLE", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"]
BIDI_LONG = {
    "Left_To_Right": "L", "Right_To_Left": "R", "Arabic_Letter": "AL", "European_Number": "EN",
    "European_Separator": "ES", "European_Terminator": "ET", "Arabic_Number": "AN",
    "Common_Separator": "CS", "Nonspacing_Mark": "NSM", "Boundary_Neutral": "BN",
    "Paragraph_Separator": "B", "Segment_Separator": "S", "White_Space": "WS", "Other_Neutral": "ON",
}
LINE_BREAK_CLASSES = ["XX", "BK", "CR", "LF", "CM", "NL", "SG", "WJ", "ZW", "GL", "SP", "ZWJ",
                      "B2", "BA", "BB", "HY", "CB", "CL", "CP", "EX", "IN", "NS", "OP", "QU", "IS",
                      "NU", "PO", "PR", "SY", "AI", "AK", "AL", "AP", "AS", "CJ", "EB", "EM", "H2",
                      "H3", "HL", "ID", "JL", "JV", "JT", "RI", "SA", "VF", "VI"]
DATA_LINE = re.compile(r"^([0-9A-F]+)(?:\.\.([0-9A-F]+))?\s*;\s*([^#\s;]+)")
MISSING_LINE = re.compile(r"^#\s*@missing:\s*([0-9A-F]+)\.\.([0-9A-F]+)\s*;\s*([^#\s;]+)")


def find(root, *names):
    for name in names:
        path = root / name
        if path.exists():
            return path
    raise FileNotFoundError(f"None of {names} found in {root}")


def read_property(path, default, normalize=lambda value: value):
    values = [default] * 0x110000
    lines = path.read_text(encoding="utf-8").splitlines()
    for line in lines:
        missing = MISSING_LINE.match(line)
        if missing:
            start, end = int(missing.group(1), 16), int(missing.group(2), 16)
            value = normalize(missing.group(3))
            for code in range(start, end + 1):
                values[code] = value
    for line in lines:
        match = DATA_LINE.match(line)
        if not match:
            continue
        start = int(match.group(1), 16)
        end = int(match.group(2), 16) if match.group(2) else start
        value = normalize(match.group(3))
        for code in range(start, end + 1):
            values[code] = value
    return values


def ranges(values, default):
    starts, result = [], []
    previous = None
    for code, value in enumerate(values):
        if value != previous:
            starts.append(code)
            result.append(value)
            previous = value
    if result and result[0] == default and starts[0] == 0:
        pass
    return starts, result


def rows(items, per_line):
    return [", ".join(items[i:i + per_line]) for i in range(0, len(items), per_line)]


def array(name, kind, items, per_line):
    out = [f"    private static readonly {kind}[] {name} = {{"]
    out += ["        " + row + "," for row in rows(items, per_line)]
    if len(out) > 1:
        out[-1] = out[-1].rstrip(",")
    out.append("    };")
    return out


def generate_combining(root):
    combining_path = find(root, "extracted/DerivedCombiningClass.txt", "DerivedCombiningClass.txt")
    properties_path = find(root, "PropList.txt")
    for path, header in ((combining_path, "# DerivedCombiningClass-16.0.0.txt"),
                         (properties_path, "# PropList-16.0.0.txt")):
        if not path.read_text(encoding="utf-8").startswith(header):
            raise ValueError(f"Expected pinned UCD 16.0.0 input: {path}")
    combining = read_property(combining_path, "0", lambda value: "0" if value == "Not_Reordered" else value)
    if any(not value.isdigit() or not 0 <= int(value) <= 255 for value in combining):
        raise ValueError("Canonical combining classes must fit their Unicode byte domain")
    starts, values = ranges(combining, "0")
    modifiers = []
    for line in properties_path.read_text(encoding="utf-8").splitlines():
        match = DATA_LINE.match(line)
        if match and match.group(3) == "Modifier_Combining_Mark":
            first = int(match.group(1), 16)
            last = int(match.group(2), 16) if match.group(2) else first
            modifiers.extend(range(first, last + 1))
    modifiers = sorted(set(modifiers))
    out = ["// <auto-generated>",
           "// Generated by Build/Generate-UnicodeTextData.py --combining-only from UCD 16.0.0.",
           "// https://www.unicode.org/Public/16.0.0/ucd/extracted/DerivedCombiningClass.txt",
           "// SHA256 " + hashlib.sha256(combining_path.read_bytes()).hexdigest(),
           "// https://www.unicode.org/Public/16.0.0/ucd/PropList.txt",
           "// SHA256 " + hashlib.sha256(properties_path.read_bytes()).hexdigest(),
           "// Unicode data files are licensed under Unicode License v3 (https://www.unicode.org/license.txt).",
           "// </auto-generated>", "", "namespace OfficeIMO.Drawing;", "",
           "internal static partial class OfficeUnicodeTextData {"]
    out += array("CombiningClassStarts", "int", [f"0x{value:X}" for value in starts], 10)
    out += array("CombiningClassValues", "byte", values, 24)
    out += array("ModifierCombiningMarks", "int", [f"0x{value:X}" for value in modifiers], 10)
    out += ["", "    /// <summary>Canonical_Combining_Class from the pinned Unicode 16.0.0 UCD.</summary>",
            "    internal static byte GetCanonicalCombiningClass(int scalar) =>",
            "        CombiningClassValues[FindRange(CombiningClassStarts, scalar)];", "",
            "    /// <summary>Modifier_Combining_Mark membership used by Unicode AMTRA.</summary>",
            "    internal static bool IsModifierCombiningMark(int scalar) =>",
            "        System.Array.BinarySearch(ModifierCombiningMarks, scalar) >= 0;", "}", ""]
    target = Path(__file__).resolve().parents[1] / "OfficeIMO.Core" / "OfficeUnicodeTextData.Combining.Generated.cs"
    target.write_text("\n".join(out), encoding="utf-8")
    print(f"Wrote {target}: {len(starts)} combining ranges, {len(modifiers)} modifier marks")


def main():
    root = Path(sys.argv[1])
    if len(sys.argv) > 2:
        if sys.argv[2:] != ["--combining-only"]:
            raise ValueError("Only --combining-only is supported")
        generate_combining(root)
        return
    bidi = read_property(find(root, "extracted/DerivedBidiClass.txt", "DerivedBidiClass.txt"), "L",
                         lambda value: BIDI_LONG.get(value, value))
    assert set(bidi) <= set(BIDI_CLASSES), set(bidi) - set(BIDI_CLASSES)
    line_break = read_property(find(root, "LineBreak.txt"), "XX")
    assert set(line_break) <= set(LINE_BREAK_CLASSES), set(line_break) - set(LINE_BREAK_CLASSES)
    east_asian = read_property(find(root, "EastAsianWidth.txt"), "N")
    wide = ["1" if value in ("F", "W", "H") else "0" for value in east_asian]
    emoji = [False] * 0x110000
    for line in find(root, "emoji/emoji-data.txt", "emoji-data.txt").read_text(encoding="utf-8").splitlines():
        match = DATA_LINE.match(line)
        if match and match.group(3) == "Extended_Pictographic":
            start = int(match.group(1), 16)
            end = int(match.group(2), 16) if match.group(2) else start
            for code in range(start, end + 1):
                emoji[code] = True
    pictographic = ["1" if value else "0" for value in emoji]

    brackets = []
    for line in find(root, "BidiBrackets.txt").read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        code, pair, kind = [field.strip() for field in line.split(";")]
        brackets.append((int(code, 16), int(pair, 16), kind))
    brackets.sort()

    bidi_starts, bidi_values = ranges(bidi, "L")
    lb_starts, lb_values = ranges(line_break, "XX")
    wide_starts, wide_values = ranges(wide, "0")
    pict_starts, pict_values = ranges(pictographic, "0")

    out = [
        "// <auto-generated>",
        "// Generated by Build/Generate-UnicodeTextData.py from the Unicode Character Database 16.0.0:",
        "// DerivedBidiClass.txt, BidiBrackets.txt, LineBreak.txt, EastAsianWidth.txt and emoji-data.txt.",
        "// Unicode data files are licensed under the Unicode License v3 (https://www.unicode.org/license.txt).",
        "// </auto-generated>",
        "",
        "namespace OfficeIMO.Drawing;",
        "",
        "/// <summary>Unicode Bidi_Class values (UAX #9, table 4).</summary>",
        "internal enum OfficeBidiClass : byte {",
        "    " + ", ".join(BIDI_CLASSES),
        "}",
        "",
        "/// <summary>Unicode Line_Break values (UAX #14, table 1).</summary>",
        "internal enum OfficeLineBreakClass : byte {",
        "    " + ", ".join(LINE_BREAK_CLASSES),
        "}",
        "",
        "/// <summary>Range tables for Unicode text layout properties.</summary>",
        "internal static partial class OfficeUnicodeTextData {",
        "    internal const string UnicodeVersion = \"16.0.0\";",
        "",
    ]
    out += array("BidiClassStarts", "int", [f"0x{value:X}" for value in bidi_starts], 10)
    out += array("BidiClassValues", "byte", [str(BIDI_CLASSES.index(value)) for value in bidi_values], 24)
    out += [""]
    out += array("LineBreakStarts", "int", [f"0x{value:X}" for value in lb_starts], 10)
    out += array("LineBreakValues", "byte", [str(LINE_BREAK_CLASSES.index(value)) for value in lb_values], 24)
    out += [""]
    out += array("EastAsianWideStarts", "int", [f"0x{value:X}" for value in wide_starts], 10)
    out += array("EastAsianWideValues", "byte", list(wide_values), 32)
    out += [""]
    out += array("ExtendedPictographicStarts", "int", [f"0x{value:X}" for value in pict_starts], 10)
    out += array("ExtendedPictographicValues", "byte", list(pict_values), 32)
    out += [""]
    out += array("BracketScalars", "int", [f"0x{code:X}" for code, _, _ in brackets], 10)
    out += array("BracketPairs", "int", [f"0x{pair:X}" for _, pair, _ in brackets], 10)
    out += array("BracketOpening", "bool", ["true" if kind == "o" else "false" for _, _, kind in brackets], 12)
    out += [
        "",
        "    internal static OfficeBidiClass GetBidiClass(int scalar) =>",
        "        (OfficeBidiClass)BidiClassValues[FindRange(BidiClassStarts, scalar)];",
        "",
        "    internal static OfficeLineBreakClass GetLineBreakClass(int scalar) =>",
        "        (OfficeLineBreakClass)LineBreakValues[FindRange(LineBreakStarts, scalar)];",
        "",
        "    /// <summary>East_Asian_Width is Fullwidth, Wide or Halfwidth.</summary>",
        "    internal static bool IsEastAsianWide(int scalar) =>",
        "        EastAsianWideValues[FindRange(EastAsianWideStarts, scalar)] != 0;",
        "",
        "    internal static bool IsExtendedPictographic(int scalar) =>",
        "        ExtendedPictographicValues[FindRange(ExtendedPictographicStarts, scalar)] != 0;",
        "",
        "    /// <summary>Gets the Bidi_Paired_Bracket of a paired bracket and whether it opens the pair.</summary>",
        "    internal static bool TryGetPairedBracket(int scalar, out int pair, out bool opening) {",
        "        int index = System.Array.BinarySearch(BracketScalars, scalar);",
        "        if (index < 0) {",
        "            pair = 0;",
        "            opening = false;",
        "            return false;",
        "        }",
        "        pair = BracketPairs[index];",
        "        opening = BracketOpening[index];",
        "        return true;",
        "    }",
        "",
        "    private static int FindRange(int[] starts, int scalar) {",
        "        if ((uint)scalar > 0x10FFFF) scalar = 0xFFFD;",
        "        int index = System.Array.BinarySearch(starts, scalar);",
        "        return index >= 0 ? index : ~index - 1;",
        "    }",
        "}",
        "",
    ]
    target = Path(__file__).resolve().parents[1] / "OfficeIMO.Core" / "OfficeUnicodeTextData.Generated.cs"
    target.write_text("\n".join(out), encoding="utf-8")
    print(f"Wrote {target}: {len(bidi_starts)} bidi, {len(lb_starts)} line-break, {len(wide_starts)} width, "
          f"{len(pict_starts)} pictographic ranges, {len(brackets)} brackets")
    generate_combining(root)


if __name__ == "__main__":
    main()
