using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    internal long RetainedProfileBytes {
        get {
            long bytes = _exif?.RetainedManagedBytes ?? 0L;
            bytes = checked(bytes + (_xmp?.LongLength ?? 0L) + (_icc?.LongLength ?? 0L) + (_iptc?.LongLength ?? 0L));
            foreach (OfficeExifValue value in _changes.Values) bytes = checked(bytes + value.RetainedByteLength + 128L);
            return bytes;
        }
    }

    private static OfficeMetadataRewriteStream CreateRewriteStream(byte[] input, OfficeImageMetadata? metadata, CancellationToken token, long additionallyRetainedBytes = 0L) {
        long profiles = metadata?.RetainedProfileBytes ?? 0L;
        // Encoding and framing can coexist with the supplied profiles and one metadata clone.
        // Reserve their transient storage before creating output; pre-size output so normal
        // image-sized writes do not double a raster-sized backing array.
        long retained = checked(input.LongLength + additionallyRetainedBytes + profiles * 4L + 65536L);
        long capacity = Math.Min(OfficeRasterGuards.MaximumEncodedBytes, checked(input.LongLength + profiles + 65536L));
        return new OfficeMetadataRewriteStream(retained, checked((int)capacity), token);
    }
}
