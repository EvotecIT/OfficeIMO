using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Rebuilds the normalized result summaries after a format adapter supplies richer structures.
    /// </summary>
    internal static OfficeDocumentReadResult EnrichDocumentResult(
        OfficeDocumentReadResult result,
        IReadOnlyList<string> capabilities,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentLink> links,
        IReadOnlyList<ReaderVisual> visuals,
        IReadOnlyList<OfficeDocumentPage> pages,
        IReadOnlyList<OfficeDocumentMetadataEntry>? formatMetadata = null) {
        if (result == null) throw new ArgumentNullException(nameof(result));

        result.CapabilitiesUsed = result.CapabilitiesUsed
            .Concat(capabilities ?? Array.Empty<string>())
            .Where(static capability => !string.IsNullOrWhiteSpace(capability))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        result.Blocks = blocks ?? Array.Empty<OfficeDocumentBlock>();
        result.Tables = tables ?? Array.Empty<ReaderTable>();
        result.Links = links ?? Array.Empty<OfficeDocumentLink>();
        result.Visuals = visuals ?? Array.Empty<ReaderVisual>();

        OfficeDocumentOcrCandidate[] ocrCandidates = BuildChunkDocumentOcrCandidates(result.Blocks, result.Assets).ToArray();
        result.OcrCandidates = ocrCandidates;
        result.Pages = AttachOcrCandidates(pages ?? Array.Empty<OfficeDocumentPage>(), ocrCandidates);
        result.Diagnostics = RefreshOcrDiagnostics(result.Diagnostics, ocrCandidates);

        IReadOnlyList<OfficeDocumentMetadataEntry> summary = BuildChunkDocumentMetadata(
            result.Kind,
            result.Chunks,
            result.Blocks,
            result.Tables,
            result.Visuals,
            result.Pages,
            result.Assets);
        result.Metadata = formatMetadata == null || formatMetadata.Count == 0
            ? summary
            : summary.Concat(formatMetadata).ToArray();
        return result;
    }

    internal static OfficeDocumentMetadataEntry BuildCountMetadataEntry(string id, string category, string name, int count) =>
        new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = count.ToString(CultureInfo.InvariantCulture),
            ValueType = "integer"
        };

    internal static IReadOnlyList<string> BuildFallbackColumns(int count) =>
        Enumerable.Range(1, Math.Max(0, count))
            .Select(static index => "Column " + index.ToString(CultureInfo.InvariantCulture))
            .ToArray();

    private static IReadOnlyList<OfficeDocumentPage> AttachOcrCandidates(
        IReadOnlyList<OfficeDocumentPage> pages,
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates) {
        for (int index = 0; index < pages.Count; index++) {
            OfficeDocumentPage page = pages[index];
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
        for (int index = 0; index < candidates.Count; index++) {
            OfficeDocumentOcrCandidate candidate = candidates[index];
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
