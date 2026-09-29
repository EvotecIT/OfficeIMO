using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfProductionPreflightReviewRegressions {
    [Fact]
    public void MissingBleedBoxUsesCropBoxForProductionNesting() {
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf("", "",
            "6 0 obj\n<< /Type /OCG /Name (Unused) >>\nendobj",
            "[6 0 R]", "[6 0 R]", pageEntries: "/CropBox [0 0 100 100] /TrimBox [0 0 150 100]");

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.InvalidPageBoxes);
        Assert.Contains(report.FixupProposals, static proposal => proposal.Box == PdfPageBoundaryBox.BleedBox);
    }

    [Fact]
    public void FullyTransparentImageDoesNotCreateResolutionFinding() {
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf(
            "/GS0 gs q 72 0 0 72 10 10 cm /Im0 Do Q",
            "/XObject << /Im0 5 0 R >> /ExtGState << /GS0 << /ca 0 >> >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Unused) >>\nendobj", "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction();

        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
    }

    [Fact]
    public void GraphicsStateFontRemainsDefiniteBesideHiddenPrintLayer() {
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf(
            "BT /GS1 gs 10 40 Td (Print) Tj ET /OC /Hidden BDC 10 10 20 20 re f EMC",
            "/ExtGState << /GS1 << /Font [5 0 R 12] >> >> /Properties << /Hidden 6 0 R >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
    }

    [Fact]
    public void SupportedPrintLayerRetainsColorAndFontFindingsBesideHiddenLayer() {
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf(
            "/OC /Visible BDC 1 0 0 rg 10 10 20 20 re f BT /F1 12 Tf 10 50 Td (Print) Tj ET EMC " +
            "/OC /Hidden BDC 0 1 0 rg 40 10 20 20 re f EMC",
            "/Font << /F1 5 0 R >> /Properties << /Visible 6 0 R /Hidden 7 0 R >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Visible) /Usage << /Print << /PrintState /ON >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R 7 0 R]", "[6 0 R 7 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
    }

    [Fact]
    public void SupportedPrintLayerRetainsImageResolutionBesideUnsupportedLayer() {
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf(
            "/OC /Visible BDC q 72 0 0 72 10 10 cm /Im0 Do Q EMC " +
            "/OC /Unsupported BDC 10 10 20 20 re f EMC",
            "/XObject << /Im0 5 0 R >> /Properties << /Unsupported 6 0 R /Visible 7 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Unsupported) /Usage << /Print << /PrintState /Maybe >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /OCG /Name (Visible) /Usage << /Print << /PrintState /ON >> >> >>\nendobj",
            "[6 0 R 7 0 R]", "[6 0 R 7 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableImageResolution);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(7)]
    public void Type3GlyphPaintRemainsDefiniteBesideHiddenLayer(int textMode) {
        const string glyph = "0 1 0 rg 0 0 20 20 re f";
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf(
            $"BT /F1 12 Tf {textMode} Tr 10 60 Td (A) Tj ET /OC /Hidden BDC 10 10 20 20 re f EMC",
            "/Font << /F1 7 0 R >> /Properties << /Hidden 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
            "/CharProcs << /A 8 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
            "/FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj\n" +
            "8 0 obj\n<< /Length " + glyph.Length + " >>\nstream\n" + glyph + "\nendstream\nendobj",
            "[6 0 R]", "[6 0 R]", size: 9);

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
    }

    [Fact]
    public void FormTextInheritsInvisibleModeInDefiniteColorFallback() {
        const string form = "BT /F1 12 Tf 1 0 0 rg 10 20 Td (Invisible) Tj ET";
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf(
            "BT 3 Tr ET /Fm Do /OC /Hidden BDC 10 10 20 20 re f EMC",
            "/Font << /F1 5 0 R >> /XObject << /Fm 7 0 R >> /Properties << /Hidden 6 0 R >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
    }

    [Fact]
    public void FormTextInheritsSelectedFontInDefiniteFontFallback() {
        const string form = "BT 10 20 Td (Form text) Tj ET";
        byte[] pdf = PdfProductionPreflightTests.RawPrintLayerPdf(
            "BT /F1 12 Tf ET /Fm Do /OC /Hidden BDC 10 10 20 20 re f EMC",
            "/Font << /F1 5 0 R >> /XObject << /Fm 7 0 R >> /Properties << /Hidden 6 0 R >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(pdf).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
    }
}
