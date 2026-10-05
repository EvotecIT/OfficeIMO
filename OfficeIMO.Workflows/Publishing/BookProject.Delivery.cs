using OfficeIMO.Epub;
using OfficeIMO.Provenance;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    /// <summary>Maximum delivery ZIP size: a bounded EPUB, package metadata, manifest and ZIP overhead.</summary>
    public const long MaximumDeliveryBytes = 134L * 1024 * 1024;
    private const long MaximumDeliveryMetadataBytes = 4L * 1024 * 1024;

    /// <summary>
    /// Exports a reviewed EPUB, its exact selected OPF payload and a versioned checksum manifest in a ZIP.
    /// This local handoff format does not perform ONIX conversion, external validation or retailer submission.
    /// Publication, OPF and manifest limits are 128 MiB, 4 MiB and 1 MiB respectively.
    /// No files are written and project history is not included.
    /// </summary>
    public byte[] ToDeliveryBytes(EpubWriteOptions? options = null,
        long maximumOutputBytes = MaximumDeliveryBytes, CancellationToken cancellationToken = default) {
        if (maximumOutputBytes <= 0 || maximumOutputBytes > MaximumDeliveryBytes)
            throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        options ??= new EpubWriteOptions();
        // Copy caller options so imposing the delivery bound does not change subsequent exports.
        var writerOptions = new EpubWriteOptions {
            CompressEntries = options.CompressEntries,
            MaxOutputBytes = Math.Min(options.MaxOutputBytes, Math.Min(MaximumPublicationBytes, maximumOutputBytes)),
            MaxExpandedBytes = options.MaxExpandedBytes, MaxEntries = options.MaxEntries,
            RemoveInvalidatedSignatures = options.RemoveInvalidatedSignatures, ModifiedAt = options.ModifiedAt
        };
        EpubWriteResult result = Export(writerOptions, cancellationToken);
        byte[] package;
        using (var input = new MemoryStream(result.Bytes, false))
        using (var archive = new ZipArchive(input, ZipArchiveMode.Read)) {
            ZipArchiveEntry entry = archive.GetEntry(_publication.PackagePath)
                ?? throw new InvalidDataException("The exported EPUB does not contain its selected package.");
            package = Read(entry, MaximumDeliveryMetadataBytes, cancellationToken);
        }
        if (_diagnostics.Count > 10_000 || result.Report.FidelityDiagnostics.Count > 10_000)
            throw new InvalidDataException("The delivery manifest exceeds its diagnostic-count bound.");
        var manifest = new BookDeliveryManifest {
            PackagePath = _publication.PackagePath,
            ImportLossAcknowledged = ImportLossAcknowledged,
            Files = [DescribeDeliveryFile("publication.epub", "application/epub+zip", result.Bytes),
                DescribeDeliveryFile("package.opf", "application/oebps-package+xml", package)],
            ImportDiagnostics = _diagnostics.Select(DescribeDeliveryDiagnostic).ToArray(),
            WriterDiagnostics = result.Report.FidelityDiagnostics.Select(DescribeDeliveryDiagnostic).ToArray()
        };
        byte[] manifestBytes;
        using (var output = new OfficeProvenanceBoundedMemoryStream(MaximumReviewBytes)) {
            JsonSerializer.Serialize(output, manifest, BookDeliveryJsonContext.Default.BookDeliveryManifest);
            manifestBytes = output.ToArray();
        }
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeProvenanceZipWriter.Write([
            Entry("publication.epub", result.Bytes, false), Entry("package.opf", package, true),
            Entry("manifest.json", manifestBytes, true)
        ], MaximumPublicationBytes + MaximumDeliveryMetadataBytes + MaximumReviewBytes,
            maximumOutputBytes: maximumOutputBytes, cancellationToken: cancellationToken);
    }

    private static BookDeliveryFile DescribeDeliveryFile(string name, string mediaType, byte[] bytes) => new() {
        Name = name, MediaType = mediaType, Bytes = bytes.LongLength, Sha256 = Convert.ToHexString(SHA256.HashData(bytes))
    };

    private static BookDeliveryDiagnostic DescribeDeliveryDiagnostic(OfficeConversionFidelityDiagnostic diagnostic) => new() {
        Code = diagnostic.Code, Message = diagnostic.Message, Source = diagnostic.Source,
        Location = diagnostic.Location, LossKind = diagnostic.LossKind.ToString()
    };
}

internal sealed class BookDeliveryManifest {
    public string Format { get; set; } = "OfficeIMO.BookDelivery";
    public int Version { get; set; } = 1;
    public string PackagePath { get; set; } = string.Empty;
    public string NativeWriterValidation { get; set; } = "passed";
    public string IndependentValidation { get; set; } = "not-performed";
    public string RetailerAcceptance { get; set; } = "not-performed";
    public bool ImportLossAcknowledged { get; set; }
    public BookDeliveryFile[] Files { get; set; } = [];
    public BookDeliveryDiagnostic[] ImportDiagnostics { get; set; } = [];
    public BookDeliveryDiagnostic[] WriterDiagnostics { get; set; } = [];
}
internal sealed class BookDeliveryFile {
    public string Name { get; set; } = string.Empty;
    public string MediaType { get; set; } = string.Empty;
    public long Bytes { get; set; }
    public string Sha256 { get; set; } = string.Empty;
}
internal sealed class BookDeliveryDiagnostic {
    public string Code { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public string? Location { get; set; }
    public string LossKind { get; set; } = string.Empty;
}
[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(BookDeliveryManifest))]
internal sealed partial class BookDeliveryJsonContext : JsonSerializerContext;
