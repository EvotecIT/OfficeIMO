using System;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    // STIX Two Math design constants, independently read by fontTools 4.46.0 from
    // macOS STIXTwoMath.otf (SHA-256 3a5f3f26f40d5698b3c62dd085d48d6663696a3f80825aab8b553d5097518e8c).
    // Glyphs are our existing synthetic rectangles; no third-party font outlines are copied.
    internal static byte[] CreateMathFont(Action<byte[]>? editTable = null) {
        int[] values = { 70, 55, 1325, 1800, 150, 258, 480, 656, 210, 368, 160, 360,
            252, 120, 230, 150, 380, 40, 135, 300, 135, 670, 470, 780, 385, 690,
            150, 300, 800, 590, 68, 68, 585, 640, 585, 640, 68, 150, 68, 68, 150,
            350, 68, 175, 68, 68, 175, 68, 68, 85, 170, 68, 78, 65, -335, 55 };
        var math = new byte[224];
        WriteUInt16(math, 0, 1);
        WriteUInt16(math, 4, 10);
        for (int index = 0; index < 4; index++) WriteUInt16(math, 10 + index * 2, unchecked((ushort)values[index]));
        for (int index = 4; index < 55; index++) WriteUInt16(math, 18 + (index - 4) * 4, unchecked((ushort)values[index]));
        WriteUInt16(math, 222, 55);
        editTable?.Invoke(math);
        return CreateFontFromCmap(CreateFormat12Cmap(new[] { 32, 50, 51, 105, 120, 121, 122 }), math: math);
    }
}
