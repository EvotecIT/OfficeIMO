using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeDeflateStreamValidator {
    /// <summary>Validates the remaining seekable stream with bounded buffering and restores its position.</summary>
    internal static bool TryValidateExact(Stream source, long maximumOutputBytes, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead || !source.CanSeek || maximumOutputBytes < 0) return false;
        long position = source.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            long count = source.Length - position;
            if (count < 0) return false;
            return TryValidateExact(new DeflateBitReader(source, count), count, maximumOutputBytes, out _, cancellationToken);
        } finally {
            source.Position = position;
        }
    }
}
