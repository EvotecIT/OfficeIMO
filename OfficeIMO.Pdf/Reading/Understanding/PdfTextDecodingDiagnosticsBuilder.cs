using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Collects one bounded diagnostic for each distinct failed font and content stream on a page.</summary>
internal sealed class PdfTextDecodingDiagnosticsBuilder {
    private readonly PdfUnderstandingPageContext _context;
    private List<PdfFontInspectionDiagnostic>? _diagnostics;
    private HashSet<(PdfFontInspectionDiagnosticCode Code, string Font, int? Stream)>? _seen;

    internal PdfTextDecodingDiagnosticsBuilder(PdfUnderstandingPageContext context) {
        _context = context;
    }

    internal PdfUnsupportedTextMappingException? FirstMappingFailure { get; private set; }
    internal long OmittedRunCount { get; private set; }

    internal void RecordUnmappableText(PdfUnsupportedTextMappingException failure, int? stream) {
        _context.ConsumeWork();
        OmittedRunCount++;
        FirstMappingFailure ??= failure;
        Add(PdfFontInspectionDiagnosticCode.UnmappableText, failure.FontResource, stream);
    }

    internal void Add(PdfFontInspectionDiagnosticCode code, string font, int? stream) {
        _seen ??= new HashSet<(PdfFontInspectionDiagnosticCode, string, int?)>();
        if (!_seen.Add((code, font, stream))) return;
        _context.ConsumeWork(2L * font.Length + 80L);
        string path = "Page " + _context.PageNumber.ToString(CultureInfo.InvariantCulture);
        if (stream is int streamNumber)
            path += "/ContentStream/" + streamNumber.ToString(CultureInfo.InvariantCulture);
        path += "/Font/" + font;
        string reason = code == PdfFontInspectionDiagnosticCode.MissingFontResource
            ? "the font resource could not be resolved"
            : "its character codes could not be mapped to Unicode";
        (_diagnostics ??= new List<PdfFontInspectionDiagnostic>()).Add(new PdfFontInspectionDiagnostic(code,
            "Text using font /" + font + " was omitted because " + reason + ".", _context.PageNumber, path));
    }

    internal IReadOnlyList<PdfFontInspectionDiagnostic> ToDiagnostics() =>
        _diagnostics is { } diagnostics ? diagnostics.AsReadOnly() : Array.Empty<PdfFontInspectionDiagnostic>();
}
