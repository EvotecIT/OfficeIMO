using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace OfficeIMO;

public static partial class OfficeDocument {
    private static void InspectSummary(OfficeDocumentInfo result, Stream source, OfficeInspectOptions settings,
        CancellationToken cancellationToken, bool sourceIsPath = false) {
        if (settings.MaximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(settings.MaximumInputBytes));
        long maximum = Math.Min(settings.MaximumInputBytes, settings.Reader.MaxInputBytes ?? long.MaxValue);
        ReaderOptions readOptions = DocumentReaderEngine.NormalizeOptions(settings.Reader, inheritProjection: false);
        readOptions.MaxInputBytes = maximum;
        var detectionOptions = settings.Detection ?? new ReaderDetectionOptions {
            Mode = readOptions.DetectionMode, MaxProbeBytes = readOptions.DetectionMaxProbeBytes,
            MaxContainerEntries = readOptions.DetectionMaxContainerEntries,
            InspectContainers = readOptions.DetectionInspectContainers
        };
        readOptions.DetectionMode = detectionOptions.Mode;
        readOptions.DetectionMaxProbeBytes = detectionOptions.MaxProbeBytes;
        readOptions.DetectionMaxContainerEntries = detectionOptions.MaxContainerEntries;
        readOptions.DetectionInspectContainers = detectionOptions.InspectContainers;
        byte[] bytes = OfficeStreamReader.ReadAllBytes(source, cancellationToken, maximum);
        using var input = new MemoryStream(bytes, writable: false);
        OfficeDocumentReader reader = settings.DocumentReader ?? new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers(settings.Formats).Build();
        result.Detection = reader.Detect(input, result.Source, detectionOptions);
        result.Format = result.Detection.Kind;
        bool encryptedPackage = result.Detection.Evidence.Contains("container:ole-encrypted-openxml-package", StringComparer.Ordinal);
        if (encryptedPackage) {
            result.IsEncrypted = true;
            if (string.IsNullOrEmpty(settings.Reader.OpenPassword)) return;
        }
        cancellationToken.ThrowIfCancellationRequested();
        input.Position = 0;
        // A supplied registry owns its custom parsing and format policy. Never bypass it with a built-in loader.
        if (settings.DocumentReader != null) {
            ApplyReadSummary(result, sourceIsPath
                ? reader.ReadDocument(result.Source, readOptions, cancellationToken)
                : reader.ReadDocument(input, result.Source, readOptions, cancellationToken), settings);
            return;
        }
        switch (result.Format) {
            case ReaderInputKind.Word:
                using (var document = OfficeIMO.Reader.Word.WordReaderAdapter.Load(input, readOptions))
                    PopulateSummary(result, document, settings, cancellationToken);
                break;
            case ReaderInputKind.Excel:
                using (var document = OfficeIMO.Reader.Excel.ExcelReaderAdapter.Load(input, result.Source, readOptions))
                    PopulateSummary(result, document, settings, cancellationToken);
                break;
            case ReaderInputKind.PowerPoint:
                using (var document = OfficeIMO.Reader.PowerPoint.PowerPointReaderAdapter.Load(input, result.Source, readOptions, cancellationToken))
                    PopulateSummary(result, document, settings, cancellationToken);
                break;
            case ReaderInputKind.Pdf:
                PdfLoadOptions pdf = SummaryPdfOptions(settings, maximum);
                PdfDocumentPreflight preflight = PdfDocument.Preflight(bytes, pdf, cancellationToken);
                ApplyPdfSummary(result, preflight, settings);
                if (settings.IncludeContent && preflight.CanReadLogicalObjects) {
                    input.Position = 0;
                    ApplyReadSummary(result, reader.ReadDocument(input, result.Source, readOptions, cancellationToken), settings);
                }
                break;
            case ReaderInputKind.Visio:
                PopulateSummary(result, OfficeIMO.Reader.Visio.VisioReaderAdapter.LoadForInspection(input, result.Source, readOptions, cancellationToken), settings, cancellationToken);
                break;
            default:
                ApplyReadSummary(result, reader.ReadDocument(input, result.Source, readOptions, cancellationToken), settings);
                break;
        }
        if (!result.IsEncrypted.HasValue && result.Detection.ContentInspected &&
            result.Detection.ContentConfidence >= ReaderDetectionConfidence.Medium &&
            !result.Detection.Evidence.Any(value => value.StartsWith("container:ole", StringComparison.Ordinal)))
            result.IsEncrypted = false;
    }

    private static PdfLoadOptions SummaryPdfOptions(OfficeInspectOptions settings, long? maximum = null) =>
        (settings.Formats.Pdf?.ReadOptions ?? settings.Formats.Pdf?.LoadOptions ?? PdfLoadOptions.UntrustedDefaults)
        .Clone(maximum, settings.Reader.OpenPassword ?? settings.Formats.Pdf?.Password);

    private static void ApplyPdfSummary(OfficeDocumentInfo result, PdfDocumentPreflight preflight, OfficeInspectOptions settings) {
        result.Format = ReaderInputKind.Pdf;
        result.PdfPreflight = preflight;
        result.IsEncrypted = preflight.Probe.HasEncryption ? true : preflight.CanRead ? false : null;
        result.PageCount = preflight.DocumentInfo?.PageCount;
        if (settings.Detailed) result.Details = (object?)preflight.DocumentInfo ?? preflight;
    }

    private static void ApplyReadSummary(OfficeDocumentInfo result, OfficeDocumentReadResult document, OfficeInspectOptions settings) {
        result.Format = document.Kind;
        // Page-like containers can represent sheets, slides or a selected subset, not source totals.
        // A normalized projection without qualified source-count evidence leaves totals unknown.
        if (settings.IncludeContent) result.Content = document;
        if (settings.Detailed && result.Details == null) result.Details = document;
    }
}
