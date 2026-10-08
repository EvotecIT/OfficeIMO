using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    private static byte[] RemoveTiffProfiles(byte[] input, OfficeImageMetadataProfileKinds kinds, CancellationToken token, out OfficeImageMetadataProfileKinds present, long additionallyRetainedBytes = 0L) {
        // Work from the last page to the first so every rewritten next-IFD link points at its final page.
        if (!OfficeTiffStructureValidator.TryValidate(input, 0, input.Length, token)) throw new FormatException("TIFF metadata is structurally invalid.");
        bool little = input[0] == 73; var pages = new List<int>(); var seen = new HashSet<int>(); int at = checked((int)OfficeExifProfileCodec.Read(input, 4, 4, little));
        while (at != 0) {
            token.ThrowIfCancellationRequested();
            if (!seen.Add(at) || pages.Count >= 1024) throw new FormatException("TIFF page chain is cyclic or too long."); pages.Add(at);
            int count = (int)OfficeExifProfileCodec.Read(input, at, 2, little); at = checked((int)OfficeExifProfileCodec.Read(input, at + 2 + count * 12, 4, little));
        }
        List<(long Start, long Length)> pixelRanges = OfficeTiffPixelRanges.Read(input, token, additionallyRetainedBytes);
        if (checked(input.LongLength * 2L + additionallyRetainedBytes + OfficeTiffPixelRanges.RetainedBytes(pixelRanges) + OfficeTiffPixelRanges.PlanningBytes) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("TIFF metadata rewriting exceeds the managed working-set limit.");
        present = OfficeImageMetadataProfileKinds.None; byte[] working = (byte[])input.Clone(); uint next = 0;
        for (int index = pages.Count - 1; index >= 0; index--) {
            token.ThrowIfCancellationRequested(); OfficeExifProfileCodec.Write(working, 4, (uint)pages[index], 4, little);
            int originalCount = (int)OfficeExifProfileCodec.Read(working, pages[index], 2, little);
            OfficeExifProfileCodec.Write(working, pages[index] + 2 + originalCount * 12, next, 4, little);
            var metadata = new OfficeImageMetadata(); ReadTiffProfiles(working, metadata, token, checked(input.LongLength + additionallyRetainedBytes + OfficeTiffPixelRanges.RetainedBytes(pixelRanges)));
            if (metadata.HasExifProfile) present |= OfficeImageMetadataProfileKinds.Exif;
            if (metadata._xmp != null) present |= OfficeImageMetadataProfileKinds.Xmp;
            if (metadata._icc != null) present |= OfficeImageMetadataProfileKinds.Icc;
            if (metadata._iptc != null) present |= OfficeImageMetadataProfileKinds.Iptc;
            if ((kinds & ~OfficeImageMetadataProfileKinds.C2pa) == OfficeImageMetadataProfileKinds.None) { next = (uint)pages[index]; continue; }
            if ((kinds & OfficeImageMetadataProfileKinds.Exif) != 0) metadata.ClearExif();
            if ((kinds & OfficeImageMetadataProfileKinds.Xmp) != 0) metadata._xmp = null;
            if ((kinds & OfficeImageMetadataProfileKinds.Icc) != 0) metadata._icc = null;
            if ((kinds & OfficeImageMetadataProfileKinds.Iptc) != 0) metadata._iptc = null;
            working = RewriteTiff(working, metadata, token, pixelRanges, preserveDensity: (kinds & OfficeImageMetadataProfileKinds.Exif) == 0, additionallyRetainedBytes: checked(input.LongLength + additionallyRetainedBytes));
            uint rewrittenIfd = (uint)OfficeExifProfileCodec.Read(working, 4, 4, little);
            int entries = (int)OfficeExifProfileCodec.Read(working, (int)rewrittenIfd, 2, little);
            OfficeExifProfileCodec.Write(working, checked((int)rewrittenIfd + 2 + entries * 12), next, 4, little); next = rewrittenIfd;
        }
        OfficeExifProfileCodec.Write(working, 4, next, 4, little);
        if (!OfficeTiffStructureValidator.TryValidate(working, 0, working.Length, token)) throw new FormatException("Rewritten TIFF metadata is structurally invalid.");
        return working;
    }
}
