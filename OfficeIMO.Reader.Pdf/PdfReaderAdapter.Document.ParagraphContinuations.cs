using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static void AddParagraphContinuationMetadata(
        List<OfficeDocumentMetadataEntry> entries,
        PdfLogicalDocument document,
        SourceMetadata source,
        IReadOnlyList<PdfLogicalPage>? selectedPages,
        PdfLogicalParagraphContinuationOptions? options) {
        PdfLogicalDocument continuationDocument = selectedPages is null
            ? document
            : document.WithPages(selectedPages);
        PdfLogicalParagraphContinuationGroup[] groups = continuationDocument
            .GetParagraphContinuationGroups(options)
            .Where(static group => group.SpansPages)
            .ToArray();
        if (groups.Length == 0) return;

        AddCountMetadata(entries, "pdf-paragraph-continuation-count", "pdf.paragraph.continuation", "Count", groups.Length);
        AddCountMetadata(
            entries,
            "pdf-paragraph-continuation-segment-count",
            "pdf.paragraph.continuation",
            "SegmentCount",
            groups.Sum(static group => group.Segments.Count));
        AddCountMetadata(
            entries,
            "pdf-paragraph-continuation-rejoined-hyphen-count",
            "pdf.paragraph.continuation",
            "RejoinedHyphenCount",
            groups.Sum(static group => group.RejoinedHyphenCount));
        AddNumberMetadata(
            entries,
            "pdf-paragraph-continuation-minimum-confidence",
            "pdf.paragraph.continuation",
            "MinimumConfidence",
            groups.Min(static group => group.Confidence));

        for (int index = 0; index < groups.Length; index++) {
            PdfLogicalParagraphContinuationGroup group = groups[index];
            string id = "pdf-paragraph-continuation-" + index.ToString("D4", CultureInfo.InvariantCulture);
            entries.Add(new OfficeDocumentMetadataEntry {
                Id = id,
                Category = "pdf.paragraph.continuation",
                Name = "ParagraphContinuation",
                Value = group.FirstPageNumber.ToString(CultureInfo.InvariantCulture) + ":" + group.LastPageNumber.ToString(CultureInfo.InvariantCulture),
                ValueType = "object",
                Location = BuildMetadataLocation(source, group.FirstPageNumber, "paragraph-continuation", id),
                Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["firstPageNumber"] = group.FirstPageNumber.ToString(CultureInfo.InvariantCulture),
                    ["lastPageNumber"] = group.LastPageNumber.ToString(CultureInfo.InvariantCulture),
                    ["segmentCount"] = group.Segments.Count.ToString(CultureInfo.InvariantCulture),
                    ["confidence"] = group.Confidence.ToString("R", CultureInfo.InvariantCulture),
                    ["evidence"] = group.Evidence.ToString(),
                    ["rejoinedHyphenCount"] = group.RejoinedHyphenCount.ToString(CultureInfo.InvariantCulture)
                }
            });
        }
    }

}
