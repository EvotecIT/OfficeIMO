using OfficeIMO.Epub;
using OfficeIMO.Provenance;
using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    /// <summary>Maximum stored project size, including its bounded EPUB and review record.</summary>
    public const long MaximumProjectBytes = 260L * 1024 * 1024;
    private const long MaximumPublicationBytes = 128L * 1024 * 1024;
    private const long MaximumReviewBytes = 1024L * 1024;
    private const string PublicationEntry = "publication.epub";
    private const string ReviewEntry = "project.json";

    /// <summary>Serializes an editable .oibook project. Publication bytes and review acceptance are retained together.</summary>
    public byte[] ToProjectBytes(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_diagnostics.Count > 10_000) throw new InvalidDataException("The project review exceeds its diagnostic-count bound.");
        byte[] publication = _publication.Write(cancellationToken: cancellationToken).Bytes;
        var record = new BookProjectRecord {
            Version = 2, Revisions = _revisions.Select(item => item.Info).ToArray(), ImportLossAcknowledged = ImportLossAcknowledged,
            Diagnostics = _diagnostics.Select(item => new BookProjectDiagnostic {
                Code = item.Code, Message = item.Message, LossKind = item.LossKind, Source = item.Source, Location = item.Location
            }).ToArray()
        };
        byte[] review = JsonSerializer.SerializeToUtf8Bytes(record, BookProjectJsonContext.Default.BookProjectRecord);
        if (review.LongLength > MaximumReviewBytes) throw new InvalidDataException("The project review exceeds its byte bound.");
        if (publication.LongLength > MaximumPublicationBytes) throw new InvalidDataException("The current publication exceeds its byte bound.");
        var storedEntries = new List<OfficeProvenanceZipWriteEntry> { Entry(PublicationEntry, publication, false), Entry(ReviewEntry, review, true) };
        storedEntries.AddRange(_revisions.Select(item => Entry(RevisionEntry(item.Info.Id), item.Bytes, false)));
        return OfficeProvenanceZipWriter.Write(storedEntries,
            MaximumPublicationBytes + MaximumReviewBytes + MaximumRevisionBytes, maximumOutputBytes: MaximumProjectBytes, cancellationToken: cancellationToken);
    }
    /// <summary>Loads an .oibook project after physical ZIP, byte, schema and entry validation. No files are extracted.</summary>
    public static BookProject LoadProject(byte[] bytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.LongLength > MaximumProjectBytes) throw new InvalidDataException("The project exceeds its input-byte bound.");
        cancellationToken.ThrowIfCancellationRequested();
        OfficeProvenanceZip.ValidateForOwningPackageMutation(bytes, new OfficeProvenanceOptions {
            MaxContainerEntries = MaximumRevisionCount + 2, MaxAssetBytes = MaximumPublicationBytes,
            MaxExpandedContainerBytes = MaximumPublicationBytes + MaximumReviewBytes + MaximumRevisionBytes, CancellationToken = cancellationToken
        }, validateOpcMetadata: false);
        using var stream = new MemoryStream(bytes, false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entries = OfficeProvenanceZip.GetValidatedEntryNames(bytes, archive);
        if (entries.Values.Distinct(StringComparer.Ordinal).Count() != entries.Count ||
            !entries.Values.Contains(PublicationEntry) || !entries.Values.Contains(ReviewEntry))
            throw new InvalidDataException("The project requires one publication and one review record with unique entries.");
        ZipArchiveEntry reviewEntry = entries.Single(item => item.Value == ReviewEntry).Key;
        ZipArchiveEntry publicationEntry = entries.Single(item => item.Value == PublicationEntry).Key;
        byte[] review = Read(reviewEntry, MaximumReviewBytes, cancellationToken);
        BookProjectRecord record = JsonSerializer.Deserialize(review, BookProjectJsonContext.Default.BookProjectRecord)
            ?? throw new InvalidDataException("The project review record is missing.");
        if (record.Version is not (1 or 2)) throw new NotSupportedException("This book project version is unsupported.");
        if (record.Diagnostics == null || record.Diagnostics.Length > 10_000 || record.Diagnostics.Any(item => item == null ||
            string.IsNullOrWhiteSpace(item.Code) || string.IsNullOrWhiteSpace(item.Source) || !Enum.IsDefined(item.LossKind)))
            throw new InvalidDataException("The project review diagnostics are invalid or exceed their count bound.");
        var diagnostics = record.Diagnostics.Select(item => new OfficeConversionFidelityDiagnostic(item.Code, item.Message,
            item.LossKind, item.Source, item.Location)).ToArray();
        byte[] publicationBytes = Read(publicationEntry, MaximumPublicationBytes, cancellationToken);
        using var publicationInput = new MemoryStream(publicationBytes, false);
        EpubPublication publication = EpubPublication.Load(publicationInput, cancellationToken: cancellationToken);
        var project = new BookProject(publication, diagnostics, record.ImportLossAcknowledged);
        LoadRevisions(project, record, entries, cancellationToken);
        return project;
    }
    private static byte[] Read(ZipArchiveEntry entry, long maximumBytes, CancellationToken token) {
        using Stream source = entry.Open();
        return OfficeWorkflowInputReader.ReadAllBytes(source, entry.FullName, maximumBytes, token);
    }
    private static OfficeProvenanceZipWriteEntry Entry(string name, byte[] bytes, bool compress) => new(name,
        bytes.LongLength, compress, new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero), 0, 0, [], [], [],
        () => new MemoryStream(bytes, false));
}

internal sealed class BookProjectRecord {
    public int Version { get; set; }
    public bool ImportLossAcknowledged { get; set; }
    public BookProjectDiagnostic[] Diagnostics { get; set; } = [];
    public BookProjectRevision[] Revisions { get; set; } = [];
}
internal sealed class BookProjectDiagnostic {
    public string Code { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public OfficeConversionLossKind LossKind { get; set; }
    public string Source { get; set; } = string.Empty;
    public string? Location { get; set; }
}
[JsonSourceGenerationOptions(UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(BookProjectRecord))]
internal sealed partial class BookProjectJsonContext : JsonSerializerContext;
