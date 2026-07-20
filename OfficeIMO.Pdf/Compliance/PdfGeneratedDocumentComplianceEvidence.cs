namespace OfficeIMO.Pdf;

internal sealed class PdfGeneratedDocumentComplianceEvidence {
    public PdfGeneratedDocumentComplianceEvidence(
        System.Collections.Generic.IReadOnlyList<PdfStandardFont> standardFonts,
        System.Collections.Generic.IReadOnlyList<PdfGeneratedFontComplianceEvidence> fontUsages,
        System.Collections.Generic.IReadOnlyList<PdfGeneratedImageAccessibilityEvidence> images,
        System.Collections.Generic.IReadOnlyList<PdfGeneratedDrawingAccessibilityEvidence> drawings,
        System.Collections.Generic.IReadOnlyList<PdfGeneratedFormAccessibilityEvidence> forms) {
        StandardFonts = standardFonts;
        FontUsages = fontUsages;
        Images = images;
        Drawings = drawings;
        Forms = forms;
    }

    public System.Collections.Generic.IReadOnlyList<PdfStandardFont> StandardFonts { get; }

    public System.Collections.Generic.IReadOnlyList<PdfGeneratedFontComplianceEvidence> FontUsages { get; }

    public System.Collections.Generic.IReadOnlyList<PdfGeneratedImageAccessibilityEvidence> Images { get; }

    public System.Collections.Generic.IReadOnlyList<PdfGeneratedDrawingAccessibilityEvidence> Drawings { get; }

    public System.Collections.Generic.IReadOnlyList<PdfGeneratedFormAccessibilityEvidence> Forms { get; }
}

internal sealed class PdfGeneratedFontComplianceEvidence {
    public PdfGeneratedFontComplianceEvidence(PdfStandardFont font, PdfOptions options) {
        Guard.StandardFont(font, nameof(font), "Generated standard-font usage contains an unsupported PDF font.");
        Guard.NotNull(options, nameof(options));
        StandardFont = font;
        Options = options;
    }

    public PdfGeneratedFontComplianceEvidence(PdfNamedFontFace font, PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        NamedFont = font;
        Options = options;
    }

    public PdfStandardFont? StandardFont { get; }

    public PdfNamedFontFace? NamedFont { get; }

    public string DisplayName => NamedFont?.FaceKey ?? StandardFont?.ToBaseFontName() ?? "unknown";

    public PdfOptions Options { get; }
}
