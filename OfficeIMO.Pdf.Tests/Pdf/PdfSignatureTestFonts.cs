using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfSignatureTestFonts {
    internal static PdfEmbeddedFont LoadAppearanceFont() => new PdfEmbeddedFont(PdfStandardFont.Helvetica,
        File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()
            ?? throw new FileNotFoundException("The bundled TrueType signature appearance fixture is required.")));
}
