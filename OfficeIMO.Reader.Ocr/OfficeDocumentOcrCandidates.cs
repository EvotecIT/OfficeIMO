namespace OfficeIMO.Reader;

// The transport permits candidates on the document, pages, or both. A page projection
// with the same local ID must identify the same candidate, including after JSON transport.
internal static class OfficeDocumentOcrCandidates {
    internal static IReadOnlyList<OfficeDocumentOcrCandidate> Collect(OfficeDocumentReadResult document) {
        var candidates = new List<OfficeDocumentOcrCandidate>();
        var byId = new Dictionary<string, OfficeDocumentOcrCandidate>(StringComparer.Ordinal);
        void Add(OfficeDocumentOcrCandidate candidate) {
            if (string.IsNullOrWhiteSpace(candidate.Id)) throw new ArgumentException("OCR candidates require a non-empty document-local ID.");
            if (byId.TryGetValue(candidate.Id, out OfficeDocumentOcrCandidate? prior)) {
                if (!SameCandidate(prior, candidate))
                    throw new ArgumentException("An OCR candidate ID identifies different source regions within one document: " + candidate.Id);
                return;
            }
            byId.Add(candidate.Id, candidate);
            candidates.Add(candidate);
        }
        foreach (OfficeDocumentOcrCandidate candidate in document.OcrCandidates ?? Array.Empty<OfficeDocumentOcrCandidate>()) Add(candidate);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            foreach (OfficeDocumentOcrCandidate candidate in page.OcrCandidates ?? Array.Empty<OfficeDocumentOcrCandidate>()) Add(candidate);
        }
        return candidates;
    }

    private static bool SameCandidate(OfficeDocumentOcrCandidate left, OfficeDocumentOcrCandidate right) =>
        ReferenceEquals(left, right) || (left.AssetId == right.AssetId && left.Kind == right.Kind
            && left.ImageCount == right.ImageCount
            && left.Location?.Path == right.Location?.Path && left.Location?.Page == right.Location?.Page
            && left.Location?.Slide == right.Location?.Slide && left.Location?.Sheet == right.Location?.Sheet
            && left.Location?.A1Range == right.Location?.A1Range && left.Location?.SourceBlockIndex == right.Location?.SourceBlockIndex
            && left.Region?.X == right.Region?.X && left.Region?.Y == right.Region?.Y
            && left.Region?.Width == right.Region?.Width && left.Region?.Height == right.Region?.Height);
}
