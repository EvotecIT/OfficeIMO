using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Square-device font heights from the optional OpenType VDMX table.</summary>
internal sealed class OfficeOpenTypeDeviceMetrics {
    private readonly Dictionary<int, (int Ascent, int Descent)> _heights;

    private OfficeOpenTypeDeviceMetrics(Dictionary<int, (int Ascent, int Descent)> heights) => _heights = heights;

    internal bool TryGetHeight(int pixelsPerEm, out double ascent, out double descent) {
        bool found = _heights.TryGetValue(pixelsPerEm, out var height);
        ascent = height.Ascent;
        descent = height.Descent;
        return found;
    }

    internal static OfficeOpenTypeDeviceMetrics? TryRead(OfficeOpenTypeReader reader) {
        // Variable fonts do not use VDMX. Invalid optional data falls back to the
        // scaled metrics, without reading into a neighboring sfnt table.
        if (reader.TryGetTable("fvar", out _, out _) ||
            !reader.TryGetTable("VDMX", out int start, out int length) || length < 6) return null;
        int version = reader.ReadUInt16(start);
        int groups = reader.ReadUInt16(start + 2);
        int ratios = reader.ReadUInt16(start + 4);
        int headerLength = 6 + ratios * 6;
        if (version > 1 || groups == 0 || ratios == 0 || headerLength > length) return null;
        for (int index = 0; index < ratios; index++) {
            int ratio = start + 6 + index * 4;
            int characterSet = reader.Data[ratio];
            // Version 0 accepts the Windows ANSI subset; version 1 accepts
            // either defined whole-font value. Ignore reserved character sets.
            if (characterSet > 1 || (version == 0 && characterSet != 1)) continue;
            int x = reader.Data[ratio + 1];
            int minimumY = reader.Data[ratio + 2];
            int maximumY = reader.Data[ratio + 3];
            if (!(x == 0 && minimumY == 0 && maximumY == 0) &&
                !(x > 0 && minimumY <= x && maximumY >= x)) continue;
            int offset = reader.ReadUInt16(start + 6 + ratios * 4 + index * 2);
            if (offset < headerLength || offset > length - 4) return null;
            int group = start + offset;
            int count = reader.ReadUInt16(group);
            if (count == 0 || count > (length - offset - 4) / 6) return null;
            var heights = new Dictionary<int, (int Ascent, int Descent)>(count);
            int previousHeight = 0;
            for (int record = 0; record < count; record++) {
                int entry = group + 4 + record * 6;
                int height = reader.ReadUInt16(entry);
                int ascent = reader.ReadInt16(entry + 2);
                int minimum = reader.ReadInt16(entry + 4);
                if (height <= previousHeight || ascent < 0 || minimum > 0 || ascent == minimum) return null;
                heights.Add(height, (ascent, -minimum));
                previousHeight = height;
            }
            return new OfficeOpenTypeDeviceMetrics(heights);
        }
        return null;
    }
}
