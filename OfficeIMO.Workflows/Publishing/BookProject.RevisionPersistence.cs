using OfficeIMO.Epub;
using System.IO.Compression;
using System.Security.Cryptography;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static string RevisionEntry(string id) => "revisions/" + id + ".epub";

    private static void LoadRevisions(BookProject project, BookProjectRecord record,
        IReadOnlyDictionary<ZipArchiveEntry, string> entries, CancellationToken token) {
        if (record.Revisions == null || record.Revisions.Length > MaximumRevisionCount ||
            (record.Version == 1 && record.Revisions.Length != 0))
            throw new InvalidDataException("The revision record is invalid or exceeds its count bound.");
        var expected = new HashSet<string>(StringComparer.Ordinal) { PublicationEntry, ReviewEntry };
        long total = 0;
        foreach (BookProjectRevision revision in record.Revisions) {
            token.ThrowIfCancellationRequested();
            if (revision == null || !Guid.TryParseExact(revision.Id, "N", out _) ||
                revision.Id != revision.Id.ToLowerInvariant() || string.IsNullOrWhiteSpace(revision.Name) || revision.Name.Length > 256 ||
                revision.CreatedUtc == default || revision.CreatedUtc.Offset != TimeSpan.Zero ||
                revision.PublicationBytes <= 0 || revision.PublicationBytes > MaximumRevisionBytes - total ||
                revision.Sha256 == null || revision.Sha256.Length != 64 || !expected.Add(RevisionEntry(revision.Id)))
                throw new InvalidDataException("The revision metadata is invalid or exceeds its byte bound.");
            ZipArchiveEntry entry = entries.SingleOrDefault(item => item.Value == RevisionEntry(revision.Id)).Key
                ?? throw new InvalidDataException("A revision publication is missing.");
            if (entry.Length != revision.PublicationBytes) throw new InvalidDataException("A revision length differs from its record.");
            byte[] bytes = Read(entry, revision.PublicationBytes, token);
            if (!string.Equals(Convert.ToHexString(SHA256.HashData(bytes)), revision.Sha256, StringComparison.Ordinal))
                throw new InvalidDataException("A revision content hash differs from its record.");
            using var input = new MemoryStream(bytes, false);
            _ = EpubPublication.Load(input, cancellationToken: token);
            total += bytes.LongLength;
            project._revisions.Add((revision, bytes));
        }
        if (!expected.SetEquals(entries.Values)) throw new InvalidDataException("The project contains undeclared entries.");
        token.ThrowIfCancellationRequested();
    }
}
