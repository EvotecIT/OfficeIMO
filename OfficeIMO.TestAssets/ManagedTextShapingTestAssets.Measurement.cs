using System;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    internal static byte[] CreateFontWithAdvance(ushort advanceWidth, params int[] scalars) {
        byte[] font = CreateFont(scalars);
        int tableCount = ReadUInt16(font, 4);
        for (int index = 0; index < tableCount; index++) {
            int record = 12 + index * 16;
            if (ReadUInt32(font, record) != 0x686D7478) continue;
            int offset = checked((int)ReadUInt32(font, record + 8));
            WriteUInt16(font, offset, advanceWidth);
            return font;
        }
        throw new InvalidOperationException("The test font is missing horizontal advances.");
    }
}
