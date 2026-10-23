using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>A font whose actual space/control contours independently determine safe nonpainting admission.</summary>
    internal static byte[] CreateControlAdmissionFont(int control = 0x200C, bool blankSpace = false,
        bool blankControl = false, bool withSubstitution = false, bool withKerning = false) {
        int[] scalars = new[] { 0x20, (int)'A', 'V', 'a', 'b', 'c', 0x0628, 0xFE8F, 0xFE90, 0xFE91, 0xFE92,
            control, 0x202C, 0x202E }.Distinct().OrderBy(value => value).ToArray();
        int Glyph(int scalar) => Array.IndexOf(scalars, scalar) + 1;
        byte[]? gsub = null;
        if (withSubstitution) {
            byte[] basis = CreateAlternatingScriptBudgetFont(1);
            int tables = ReadUInt16(basis, 4);
            for (int index = 0; index < tables; index++) {
                int record = 12 + index * 16;
                if (ReadUInt32(basis, record) != 0x47535542) continue;
                gsub = new byte[checked((int)ReadUInt32(basis, record + 12))];
                Array.Copy(basis, checked((int)ReadUInt32(basis, record + 8)), gsub, 0, gsub.Length);
                int lookups = ReadUInt16(gsub, 8), lookup = lookups + ReadUInt16(gsub, lookups + 2);
                int subtable = lookup + ReadUInt16(gsub, lookup + 6), coverage = subtable + ReadUInt16(gsub, subtable + 2);
                WriteUInt16(gsub, subtable + 6, checked((ushort)Glyph(0x0628)));
                WriteUInt16(gsub, coverage + 4, checked((ushort)Glyph(0x0628)));
                break;
            }
            if (gsub == null) throw new InvalidOperationException("The fixture GSUB table is unavailable.");
        }
        var blanks = new List<int>();
        if (blankSpace) blanks.Add(Glyph(0x20));
        if (blankControl) blanks.Add(Glyph(control));
        byte[]? kern = withKerning ? CreateKernTable(checked((ushort)Glyph('A')), checked((ushort)Glyph('V')), -120) : null;
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(scalars), glyphCount: scalars.Length + 1,
            gsub: gsub, kern: kern, blankGlyphIds: blanks);
    }
}
