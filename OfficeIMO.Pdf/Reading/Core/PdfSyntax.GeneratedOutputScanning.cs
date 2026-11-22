namespace OfficeIMO.Pdf;

internal readonly struct PdfStructuralMarkerCounts {
    internal PdfStructuralMarkerCounts(int indirectObjectHeaders, int startXrefMarkers, int maximumObjectCharacters, int maximumRawStreamBytes) {
        IndirectObjectHeaders = indirectObjectHeaders;
        StartXrefMarkers = startXrefMarkers;
        MaximumObjectCharacters = maximumObjectCharacters;
        MaximumRawStreamBytes = maximumRawStreamBytes;
    }

    internal int IndirectObjectHeaders { get; }
    internal int StartXrefMarkers { get; }
    internal int MaximumObjectCharacters { get; }
    internal int MaximumRawStreamBytes { get; }
}

internal static partial class PdfSyntax {
    internal static int CountIndirectObjectHeaders(byte[] pdf, PdfReadLimits limits) =>
        InspectStructuralMarkers(pdf, limits, measureObjects: false).IndirectObjectHeaders;

    internal static PdfStructuralMarkerCounts InspectStructuralMarkers(byte[] pdf, PdfReadLimits limits) =>
        InspectStructuralMarkers(pdf, limits, measureObjects: true);

    /// <summary>
    /// Counts only the indirect-object headers and startxref markers. A source document's counts are compared with its
    /// generated outputs; their object and stream sizes are not, so the object walk is skipped and they read zero.
    /// </summary>
    internal static PdfStructuralMarkerCounts CountSourceStructuralMarkers(ReadOnlyMemory<byte> pdf, PdfReadLimits limits) =>
        InspectStructuralMarkers(pdf, limits, measureObjects: false);

    private static PdfStructuralMarkerCounts InspectStructuralMarkers(ReadOnlyMemory<byte> pdf, PdfReadLimits limits, bool measureObjects) {
        Guard.NotNull(limits, nameof(limits));
        limits.Validate();
        var parseTimer = System.Diagnostics.Stopwatch.StartNew();
        var text = new PdfByteContentText(pdf);
        int count = 0;
        int maximumObjectCharacters = 0;
        int maximumRawStreamBytes = 0;
        int cursor = 0;
        while (TryFindIndirectObjectHeader(
            text,
            cursor,
            text.Length,
            out IndirectObjectHeader header,
            parseTimer,
            limits)) {
            count = checked(count + 1);
            cursor = header.Index + header.Length;
        }

        int objectCursor = measureObjects ? 0 : text.Length;
        while (TryFindIndirectObjectHeader(
            text,
            objectCursor,
            text.Length,
            out IndirectObjectHeader header,
            parseTimer,
            limits)) {
            int bodyStart = header.Index + header.Length;
            int objectEnd = FindObjectEnd(text, bodyStart, out int rawStreamBytes);
            if (objectEnd < bodyStart) {
                objectCursor = bodyStart;
                continue;
            }

            int bodyEnd = objectEnd >= 6 &&
                MatchesAt(text, objectEnd - 6, "endobj")
                    ? objectEnd - 6
                    : objectEnd;
            maximumObjectCharacters = Math.Max(maximumObjectCharacters, Math.Max(0, bodyEnd - bodyStart));
            maximumRawStreamBytes = Math.Max(maximumRawStreamBytes, rawStreamBytes);
            objectCursor = objectEnd;
        }

        int startXrefMarkers = 0;
        int startXrefCursor = 0;
        while (TryReadNextStartXrefOffset(pdf.Span, ref startXrefCursor, out _)) {
            startXrefMarkers = checked(startXrefMarkers + 1);
        }

        ThrowIfParsingTimeExceeded(parseTimer, limits);
        return new PdfStructuralMarkerCounts(count, startXrefMarkers, maximumObjectCharacters, maximumRawStreamBytes);
    }

}
