using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    internal static byte[] CreateFontWithDeviceMetrics(byte[]? deviceMetrics = null) {
        if (deviceMetrics == null) {
            deviceMetrics = new byte[28];
            WriteUInt16(deviceMetrics, 0, 1);
            WriteUInt16(deviceMetrics, 2, 1);
            WriteUInt16(deviceMetrics, 4, 1);
            deviceMetrics[6] = deviceMetrics[7] = deviceMetrics[8] = deviceMetrics[9] = 1;
            WriteUInt16(deviceMetrics, 10, 12);
            WriteUInt16(deviceMetrics, 12, 2);
            deviceMetrics[14] = 67; deviceMetrics[15] = 133;
            WriteUInt16(deviceMetrics, 16, 67);
            WriteUInt16(deviceMetrics, 18, 60);
            WriteInt16(deviceMetrics, 20, -15);
            WriteUInt16(deviceMetrics, 22, 133);
            WriteUInt16(deviceMetrics, 24, 119);
            WriteInt16(deviceMetrics, 26, -30);
        }
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)'A', (int)'B' }),
            glyphCount: 3, vdmx: deviceMetrics);
    }

    internal static byte[] CreateFontWithUniformAdvance(ushort unitsPerEm, ushort advanceWidth, params int[] scalars) {
        if (scalars == null || scalars.Length == 0) throw new ArgumentException("At least one scalar is required.", nameof(scalars));
        int[] ordered = new SortedSet<int>(scalars).ToArray();
        byte[] font = CreateFontFromCmap(CreateDistinctFormat12Cmap(ordered),
            glyphCount: ordered.Length + 1, includeTrailingMetric: true);
        return WithHorizontalMetrics(font, unitsPerEm, advanceWidth);
    }

    internal static byte[] CreateFontWithFontUnitKerning(ushort unitsPerEm, int left, int right, short adjustment) =>
        WithHorizontalMetrics(CreateFontFromCmap(CreateFormat12Cmap(left, 1, right, 2),
            glyphCount: 3, includeTrailingMetric: true, kern: CreateKernTable(1, 2, adjustment)), unitsPerEm, 500);

    private static byte[] WithHorizontalMetrics(byte[] font, ushort unitsPerEm, ushort advanceWidth) {
        int count = ReadUInt16(font, 4);
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            int offset = checked((int)ReadUInt32(font, record + 8));
            uint tag = ReadUInt32(font, record);
            if (tag == 0x68656164) WriteUInt16(font, offset + 18, unitsPerEm);
            // These fixtures have one long horizontal metric; all other glyphs
            // reuse its advance and supply only their left side bearings.
            if (tag == 0x686D7478) WriteUInt16(font, offset, advanceWidth);
        }
        return font;
    }
}
