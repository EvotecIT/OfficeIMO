using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Archives;

public sealed partial class ArchiveInspection {
    /// <summary>Copies one complete verified member to a writable caller-owned stream at its current position.
    /// ZIP members spool to a private temporary file and are fully verified before the destination receives bytes;
    /// CAB/MSI retain their complete in-memory member-read behavior. Integrity failures leave the destination untouched.
    /// Cancellation or I/O failure during the final copy can leave a verified prefix; the caller owns output rollback.
    /// The destination is left open. This operation does not make input snapshots lazy or bound process peak memory.</summary>
    public void CopyTo(int ordinal, Stream destination, CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("A writable destination is required.", nameof(destination));
        lock (_gate) {
            ThrowIfDisposed(); cancellationToken.ThrowIfCancellationRequested();
            if (ordinal < 0 || ordinal >= Entries.Count) throw new ArgumentOutOfRangeException(nameof(ordinal));
            var entry = Entries[ordinal];
            if (entry.Diagnostic != null) throw new NotSupportedException(entry.Diagnostic.Code + ": " + entry.Diagnostic.Message);
            if (Format != ArchiveFormat.Zip) {
                using var member = OpenRead(ordinal, cancellationToken);
                Copy(member, destination, cancellationToken);
                return;
            }
            using var spool = ArchiveTemporaryFile.Create();
            VerifyZipMember(ordinal, spool, cancellationToken);
            spool.Position = 0;
            Copy(spool, destination, cancellationToken);
        }
    }

    private static void Copy(Stream source, Stream destination, CancellationToken token) {
        var buffer = new byte[65536];
        while (true) {
            token.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            destination.Write(buffer, 0, read);
        }
    }
}
