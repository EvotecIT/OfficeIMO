using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static OfficeDocumentReadResult EnrichBuiltInDocumentResult(
        string path,
        ReaderInputKind kind,
        ReaderOptions options,
        OfficeDocumentReadResult result,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (kind) {
            case ReaderInputKind.Word:
                using (WordDocument document = LoadWordForReader(path,
                           options)) {
                    return ApplyWordRichMapping(document.CreateInspectionSnapshot(), options, result);
                }
            case ReaderInputKind.Excel:
                using (ExcelDocument document = LoadExcelForRichMapping(path, options)) {
                    return ApplyExcelRichMapping(document.CreateInspectionSnapshot(), options, result);
                }
            case ReaderInputKind.PowerPoint:
                using (PowerPointPresentation presentation =
                       LoadPowerPointForReader(path, options)) {
                    return ApplyPowerPointRichMapping(presentation, options, result, cancellationToken);
                }
            default:
                return result;
        }
    }

    private static OfficeDocumentReadResult EnrichBuiltInDocumentResult(
        Stream stream,
        string sourceName,
        ReaderInputKind kind,
        ReaderOptions options,
        OfficeDocumentReadResult result,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (kind) {
            case ReaderInputKind.Word:
                using (WordDocument document = LoadWordForReader(stream,
                           options)) {
                    return ApplyWordRichMapping(document.CreateInspectionSnapshot(), options, result);
                }
            case ReaderInputKind.Excel:
                stream.Position = 0;
                using (ExcelDocument document = LoadExcelForRichMapping(stream, sourceName, options)) {
                    return ApplyExcelRichMapping(document.CreateInspectionSnapshot(), options, result);
                }
            case ReaderInputKind.PowerPoint:
                stream.Position = 0;
                using (PowerPointPresentation presentation =
                       LoadPowerPointForReader(stream, options)) {
                    return ApplyPowerPointRichMapping(presentation, options, result, cancellationToken);
                }
            default:
                return result;
        }
    }

    private static ExcelDocument LoadExcelForRichMapping(string path, ReaderOptions options) {
        return IsLegacyExcelExtension(path)
            ? LoadLegacyExcelForReader(path, options)
            : LoadOpenXmlExcelForReader(path, options);
    }

    private static ExcelDocument LoadExcelForRichMapping(Stream stream, string sourceName, ReaderOptions options) {
        return IsLegacyExcelExtension(sourceName)
            ? LoadLegacyExcelForReader(stream, options)
            : LoadOpenXmlExcelForReader(stream, options);
    }

    private static OfficeDocumentReadResult FinalizeRichMapping(
        OfficeDocumentReadResult result,
        IReadOnlyList<string> capabilities,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentLink> links,
        IReadOnlyList<ReaderVisual> visuals,
        IReadOnlyList<OfficeDocumentPage> pages,
        IReadOnlyList<OfficeDocumentMetadataEntry>? formatMetadata = null) {
        result.CapabilitiesUsed = result.CapabilitiesUsed
            .Concat(capabilities)
            .Where(static capability => !string.IsNullOrWhiteSpace(capability))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        result.Blocks = blocks;
        result.Tables = tables;
        result.Links = links;
        result.Visuals = visuals;

        OfficeDocumentOcrCandidate[] ocrCandidates = BuildChunkDocumentOcrCandidates(blocks, result.Assets).ToArray();
        result.OcrCandidates = ocrCandidates;
        result.Pages = AttachOcrCandidates(pages, ocrCandidates);
        result.Diagnostics = RefreshOcrDiagnostics(result.Diagnostics, ocrCandidates);

        IReadOnlyList<OfficeDocumentMetadataEntry> summary = BuildChunkDocumentMetadata(
            result.Kind,
            result.Chunks,
            blocks,
            tables,
            visuals,
            result.Pages,
            result.Assets);
        result.Metadata = formatMetadata == null || formatMetadata.Count == 0
            ? summary
            : summary.Concat(formatMetadata).ToArray();
        return result;
    }

    private static IReadOnlyList<OfficeDocumentPage> AttachOcrCandidates(
        IReadOnlyList<OfficeDocumentPage> pages,
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates) {
        for (int i = 0; i < pages.Count; i++) {
            OfficeDocumentPage page = pages[i];
            page.OcrCandidates = candidates.Where(candidate => IsSameContainer(candidate.Location, page.Location)).ToArray();
        }
        return pages;
    }

    private static bool IsSameContainer(ReaderLocation left, ReaderLocation right) {
        if (right.Page.HasValue) return left.Page == right.Page;
        if (right.Slide.HasValue) return left.Slide == right.Slide;
        if (!string.IsNullOrWhiteSpace(right.Sheet)) return string.Equals(left.Sheet, right.Sheet, StringComparison.Ordinal);
        return false;
    }

    private static IReadOnlyList<OfficeDocumentDiagnostic> RefreshOcrDiagnostics(
        IReadOnlyList<OfficeDocumentDiagnostic> existing,
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates) {
        var diagnostics = existing
            .Where(static diagnostic => diagnostic.Category != OfficeDocumentDiagnosticCategory.Ocr)
            .ToList();
        for (int i = 0; i < candidates.Count; i++) {
            OfficeDocumentOcrCandidate candidate = candidates[i];
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Ocr,
                Code = "ocr-needed",
                Message = candidate.Reason ?? "OCR may be needed before text extraction is complete.",
                Source = "officeimo.reader",
                IsRecoverable = true,
                Location = candidate.Location
            });
        }
        return diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : diagnostics.ToArray();
    }
}
