using System;
using System.Collections.Generic;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>Rectangular outlines and explicit connector records; no third-party outlines.</summary>
    internal static byte[] CreateMathConstructionFont(Action<byte[]>? edit = null, bool overhangs = false) {
        var bytes = new List<byte>();
        int Allocate(int count) { int start = bytes.Count; for (int i = 0; i < count; i++) bytes.Add(0); return start; }
        void Put(int p, int value) { bytes[p] = (byte)(value >> 8); bytes[p + 1] = (byte)value; }
        int Coverage(params int[] glyphs) { int p = Allocate(4 + glyphs.Length * 2); Put(p, 1); Put(p + 2, glyphs.Length); for (int i = 0; i < glyphs.Length; i++) Put(p + 4 + i * 2, glyphs[i]); return p; }
        Allocate(224); Put(0, 1); Put(4, 10);
        int[] constants = { 70, 55, 1325, 1800, 150, 258, 480, 656, 210, 368, 160, 360,
            252, 120, 230, 150, 380, 40, 135, 300, 135, 670, 470, 780, 385, 690,
            150, 300, 800, 590, 68, 68, 585, 640, 585, 640, 68, 150, 68, 68, 150,
            350, 68, 175, 68, 68, 175, 68, 68, 85, 170, 68, 78, 65, -335, 55 };
        for (int i = 0; i < 4; i++) Put(10 + i * 2, constants[i]);
        for (int i = 4; i < 55; i++) Put(18 + (i - 4) * 4, constants[i]); Put(222, 55);
        int info = Allocate(8); Put(6, info);
        int italic = Allocate(12); Put(info, italic - info); Put(italic + 2, 2);
        Put(italic + 4, 150); Put(italic + 8, 200);
        Put(italic, Coverage(6, 8) - italic);
        int accents = Allocate(12); Put(info + 2, accents - info); Put(accents + 2, 2);
        Put(accents + 4, 100); Put(accents + 8, 350); Put(accents, Coverage(4, 6) - accents);
        int kern = Allocate(12); Put(info + 6, kern - info); Put(kern + 2, 1); Put(kern, Coverage(6) - kern);
        int tr = Allocate(14); Put(kern + 4, tr - kern); Put(tr, 1); Put(tr + 2, 600); Put(tr + 6, -50); Put(tr + 10, 120);
        int br = Allocate(6); Put(kern + 8, br - kern); Put(br + 2, -100);
        int variants = Allocate(20); Put(8, variants); Put(variants, 100); Put(variants + 6, 4); Put(variants + 8, 1);
        Put(variants + 2, Coverage(1, 2, 8, 9) - variants);
        for (int i = 0; i < 4; i++) {
            int construction = Allocate(12); Put(variants + 10 + i * 2, construction - variants);
            Put(construction + 2, 2); Put(construction + 4, 10); Put(construction + 6, 1400);
            Put(construction + 8, 11); Put(construction + 10, 2200);
            int assembly = Allocate(36); Put(construction, assembly - construction); Put(assembly, 220); Put(assembly + 4, 3);
            // bottom, repeated middle, top; 100..200 design units of legal overlap.
            for (int j = 0; j < 3; j++) {
                int part = assembly + 6 + j * 10; Put(part, 12 + j);
                Put(part + 2, j == 0 ? 0 : 200); Put(part + 4, j == 2 ? 0 : 200);
                Put(part + 6, j == 1 ? 600 : 700); Put(part + 8, j == 1 ? 1 : 0);
            }
        }
        int horizontalCoverage = Allocate(10); Put(horizontalCoverage, 2); Put(horizontalCoverage + 2, 1);
        Put(horizontalCoverage + 4, 4); Put(horizontalCoverage + 6, 4); // singleton range, index zero.
        Put(variants + 4, horizontalCoverage - variants);
        int horizontal = Allocate(8); Put(variants + 18, horizontal - variants);
        Put(horizontal + 2, 1); Put(horizontal + 4, 15); Put(horizontal + 6, 1200);
        int horizontalAssembly = Allocate(36); Put(horizontal, horizontalAssembly - horizontal); Put(horizontalAssembly + 4, 3);
        for (int i = 0; i < 3; i++) {
            int part = horizontalAssembly + 6 + i * 10; Put(part, 16 + i);
            Put(part + 2, i == 0 ? 0 : 200); Put(part + 4, i == 2 ? 0 : 200);
            Put(part + 6, i == 1 ? 600 : 700); Put(part + 8, i == 1 ? 1 : 0);
        }
        byte[] math = bytes.ToArray(); edit?.Invoke(math);
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { 40, 41, 50, 94, 105, 120, 121, 0x2211, 0x221A }),
            glyphCount: 19, math: math,
            glyphHeights: new Dictionary<int, int> { [4] = 200, [5] = 700, [6] = 700, [10] = 1400, [11] = 2200, [12] = 700, [13] = 600, [14] = 700,
                [15] = 200, [16] = 200, [17] = 200, [18] = 200 },
            glyphBottoms: new Dictionary<int, int> { [4] = 500, [15] = 500, [16] = 500, [17] = 500, [18] = 500 },
            glyphWidths: new Dictionary<int, int> { [5] = overhangs ? 800 : 400, [15] = 1200, [16] = 700, [17] = 600, [18] = 700 },
            glyphLefts: overhangs ? new Dictionary<int, int> { [6] = -100 } : null);
    }
}
