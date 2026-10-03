using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Email.Store;

internal sealed partial class MailboxDirectoryStoreSessionBackend {
    internal string GetCatalogFingerprint(CancellationToken cancellationToken) {
        ValidateCatalog(cancellationToken);
        using IncrementalHash fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        AppendFingerprint(fingerprint, "OfficeIMO.MailboxDirectory.Catalog.v3");
        AppendFolders(fingerprint);
        AppendInt64(fingerprint, _files.Count);
        foreach (MailboxFile file in _files) {
            cancellationToken.ThrowIfCancellationRequested();
            var info = new FileInfo(file.Path);
            AppendFingerprint(fingerprint, file.RelativePath);
            try {
                info.Refresh();
                AppendFingerprint(fingerprint, info.Exists
                    ? info.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : "missing");
                AppendFingerprint(fingerprint, info.Exists
                    ? info.LastWriteTimeUtc.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : "missing");
                AppendFingerprint(fingerprint, info.Exists &&
                    (info.Attributes & FileAttributes.ReparsePoint) != 0
                        ? "reparse"
                        : "regular");
            } catch (Exception exception) when (
                exception is IOException || exception is UnauthorizedAccessException) {
                AppendFingerprint(fingerprint, "unavailable");
            }
        }
        ValidateCatalog(cancellationToken);
        return EmailHashing.ToHexLower(fingerprint.GetHashAndReset());
    }

    internal string GetContentFingerprint(CancellationToken cancellationToken) {
        ValidateCatalog(cancellationToken);
        using IncrementalHash fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        AppendFingerprint(fingerprint, "OfficeIMO.MailboxDirectory.Content.v3");
        AppendFolders(fingerprint);
        AppendInt64(fingerprint, _files.Count);
        var buffer = new byte[64 * 1024];
        long aggregateLength = 0;
        foreach (MailboxFile file in _files) {
            cancellationToken.ThrowIfCancellationRequested();
            AppendFingerprint(fingerprint, file.RelativePath);
            var info = new FileInfo(file.Path);
            info.Refresh();
            if (!info.Exists || (info.Attributes & FileAttributes.ReparsePoint) != 0) {
                throw new InvalidDataException("A mailbox-directory source changed after it was indexed.");
            }
            using (FileStream stream = OpenRegularMailboxFile(file.Path)) {
                long declaredLength = stream.Length;
                aggregateLength = AddBounded(aggregateLength, declaredLength);
                AppendInt64(fingerprint, declaredLength);
                long totalRead = 0;
                while (totalRead < declaredLength) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = stream.Read(buffer, 0,
                        (int)Math.Min(buffer.Length, declaredLength - totalRead));
                    if (read == 0) {
                        throw new InvalidDataException("A mailbox-directory source changed while it was fingerprinted.");
                    }
                    fingerprint.AppendData(buffer, 0, read);
                    totalRead += read;
                }
                if (totalRead != declaredLength || stream.Length != declaredLength) {
                    throw new InvalidDataException("A mailbox-directory source changed while it was fingerprinted.");
                }
            }
        }
        if (aggregateLength != _sourceLength) {
            throw new InvalidDataException("The mailbox-directory aggregate source length changed after it was indexed.");
        }
        ValidateCatalog(cancellationToken);
        return EmailHashing.ToHexLower(fingerprint.GetHashAndReset());
    }

    private static void AppendFingerprint(IncrementalHash fingerprint, string value) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        AppendInt64(fingerprint, bytes.Length);
        fingerprint.AppendData(bytes);
    }

    private void AppendFolders(IncrementalHash fingerprint) {
        AppendInt64(fingerprint, _folders.Count);
        foreach (EmailStoreFolderInfo folder in _folders) AppendFingerprint(fingerprint, folder.Id);
    }

    private static void AppendInt64(IncrementalHash fingerprint, long value) {
        var bytes = new byte[8];
        ulong unsigned = unchecked((ulong)value);
        for (int index = 0; index < bytes.Length; index++) {
            bytes[index] = (byte)(unsigned >> (index * 8));
        }
        fingerprint.AppendData(bytes);
    }

}
