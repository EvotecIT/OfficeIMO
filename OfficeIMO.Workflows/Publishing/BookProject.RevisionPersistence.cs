using OfficeIMO.Internal;
using OfficeIMO.Epub;
using System.IO.Compression;
using System.Security.Cryptography;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static string RevisionEntry(string id) => "revisions/" + id + ".epub";

    private static long CountPublicationExpandedBytes(byte[] bytes, long remaining, CancellationToken token) {
        if (remaining < 0) throw new InvalidDataException("The project revision expansion bound is exhausted.");
        using var archive = new ZipArchive(new MemoryStream(bytes, false), ZipArchiveMode.Read);
        long expanded = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            token.ThrowIfCancellationRequested();
            if (entry.Length > remaining - expanded)
                throw new InvalidDataException("The project revision expansion bound is exhausted.");
            expanded += entry.Length;
        }
        return expanded;
    }

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
                revision.Id != revision.Id.ToLowerInvariant() || (revision.Name == null || string.IsNullOrWhiteSpace(revision.Name)) || revision.Name.Length > 256 ||
                revision.CreatedUtc == default || revision.CreatedUtc.Offset != TimeSpan.Zero ||
                revision.PublicationBytes <= 0 || revision.PublicationBytes > MaximumRevisionBytes - total ||
                revision.Sha256 == null || revision.Sha256.Length != 64 || !expected.Add(RevisionEntry(revision.Id)))
                throw new InvalidDataException("The revision metadata is invalid or exceeds its byte bound.");
            ZipArchiveEntry entry = entries.SingleOrDefault(item => item.Value == RevisionEntry(revision.Id)).Key
                ?? throw new InvalidDataException("A revision publication is missing.");
            if (entry.Length != revision.PublicationBytes) throw new InvalidDataException("A revision length differs from its record.");
            byte[] bytes = Read(entry, revision.PublicationBytes, token);
            if (!string.Equals(OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(bytes)), revision.Sha256, StringComparison.Ordinal))
                throw new InvalidDataException("A revision content hash differs from its record.");
            long remaining = MaximumRevisionExpandedBytes - project._revisionExpandedBytes;
            long expanded = CountPublicationExpandedBytes(bytes, remaining, token);
            var loadOptions = new EpubPublicationLoadOptions();
            loadOptions.MaxExpandedBytes = Math.Min(loadOptions.MaxExpandedBytes, remaining);
            using var input = new MemoryStream(bytes, false);
            _ = EpubPublication.Load(input, loadOptions, token);
            total += bytes.LongLength;
            project._revisionExpandedBytes += expanded;
            project._revisions.Add((revision, bytes, expanded));
        }
        if (!expected.SetEquals(entries.Values)) throw new InvalidDataException("The project contains undeclared entries.");
        token.ThrowIfCancellationRequested();
    }
}
