namespace OfficeIMO.Email.Store;

/// <summary>Detects and reads supported mailbox-store artifacts without native or third-party parser dependencies.</summary>
public sealed class EmailStoreReader {
    private readonly EmailStoreReaderOptions _options;

    /// <summary>Creates a reader with bounded defaults.</summary>
    public EmailStoreReader(EmailStoreReaderOptions? options = null) {
        _options = options ?? EmailStoreReaderOptions.Default;
    }

    /// <summary>Detects a store format from a bounded header.</summary>
    public static EmailStoreFormat DetectFormat(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (Directory.Exists(path)) return EmailStoreFormat.MailboxDirectory;
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
            return DetectFormat(stream, Path.GetFileName(path));
        }
    }

    /// <summary>Detects a store format without changing a seekable stream's position.</summary>
    public static EmailStoreFormat DetectFormat(Stream stream, string? sourceName = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The source stream must be readable.", nameof(stream));

        long position = stream.CanSeek ? stream.Position : 0;
        var header = new byte[24];
        int count = 0;
        try {
            if (stream.CanSeek) stream.Position = 0;
            while (count < header.Length) {
                int read = stream.Read(header, count, header.Length - count);
                if (read == 0) break;
                count += read;
            }
        } finally {
            if (stream.CanSeek) stream.Position = position;
        }

        if (count >= 12 && header[0] == 0x21 && header[1] == 0x42 && header[2] == 0x44 && header[3] == 0x4E) {
            if (header[8] == 0x53 && header[9] == 0x4F) return EmailStoreFormat.Ost;
            if (header[8] == 0x53 && header[9] == 0x4D) return EmailStoreFormat.Pst;
        }

        string extension = Path.GetExtension(sourceName ?? string.Empty);
        if (string.Equals(extension, ".olm", StringComparison.OrdinalIgnoreCase) &&
            count >= 4 && header[0] == 0x50 && header[1] == 0x4B) return EmailStoreFormat.Olm;
        if (string.Equals(extension, ".emlx", StringComparison.OrdinalIgnoreCase)) return EmailStoreFormat.Emlx;
        int mboxOffset = count >= 3 && header[0] == 0xEF && header[1] == 0xBB && header[2] == 0xBF
            ? 3
            : 0;
        if (string.Equals(extension, ".mbox", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".mbx", StringComparison.OrdinalIgnoreCase) ||
            count - mboxOffset >= 5 && header[mboxOffset] == (byte)'F' &&
            header[mboxOffset + 1] == (byte)'r' && header[mboxOffset + 2] == (byte)'o' &&
            header[mboxOffset + 3] == (byte)'m' && header[mboxOffset + 4] == (byte)' ') {
            return EmailStoreFormat.Mbox;
        }
        return EmailStoreFormat.Unknown;
    }

    /// <summary>Reads a file while keeping random-access parsing off the large-object heap.</summary>
    public EmailStoreReadResult Read(string path, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        using (EmailStoreSession session = EmailStoreSession.Open(path, _options, cancellationToken)) {
            return session.ReadAll(cancellationToken);
        }
    }

    /// <summary>Reads a seekable stream and leaves it open.</summary>
    public EmailStoreReadResult Read(Stream stream, string? sourceName = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead || !stream.CanSeek) {
            throw new ArgumentException("Email-store streams must be readable and seekable.", nameof(stream));
        }
        using (EmailStoreSession session = EmailStoreSession.Open(
            stream, sourceName, _options, leaveOpen: true, cancellationToken)) {
            return session.ReadAll(cancellationToken);
        }
    }
}
