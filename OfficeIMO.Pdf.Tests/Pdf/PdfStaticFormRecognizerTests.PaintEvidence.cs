using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfStaticFormRecognizerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void VisibleTextInsideFieldReportsOccupancyIncludingArtifacts(bool artifact) {
        string value = "BT /F1 12 Tf 100 85 Td (Ada) Tj ET";
        if (artifact) value = "/Artifact BMC " + value + " EMC";
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S " + value,
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void ArtifactTextOutsideFieldDoesNotBecomeALabel() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S /Artifact BMC BT /F1 12 Tf 10 85 Td (Name) Tj ET EMC",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Theory]
    [InlineData(3)]
    [InlineData(81)]
    public void OcrTextInsideFieldReportsOccupancy(int textLength) {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S");
        var labels = new[] {
            new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D),
            new PdfStaticFormTextEvidence(1, new string('A', textLength), 100D, 103D, 130D, 117D, 1D)
        };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Theory]
    [InlineData("0.5")]
    [InlineData("0")]
    public void ThinStrokeInsideCheckboxReportsOccupancy(string strokeWidth) {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 15 15 re S " + strokeWidth + " w 82 88 m 92 88 l S");
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Theory]
    [InlineData("20 w 80 80 120 20 re S")]
    [InlineData("20 w 80 80 15 15 re S")]
    [InlineData("40 w 80 100 m 200 100 l S")]
    public void StrokeFillingTheWritingAreaIsNotAnEmptyField(string paint) {
        byte[] source = StaticPdf("0 0 0 RG " + paint);
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 90D, 70D, 120D, 1D) };

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label).Proposals);
    }

    [Fact]
    public void DarkNativeLabelOnMatchingOpaqueBackdropIsNotEvidence() {
        OfficeShape backdrop = OfficeShape.Rectangle(90D, 35D);
        backdrop.FillColor = OfficeColor.Black;
        backdrop.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Shape(backdrop, 10D, 20D)
                .Text("Name:", 20D, 28D, 70D, 20D, color: PdfColor.FromRgb(0, 0, 0))
                .Shape(Box(140D, 20D), 100D, 28D)).ToBytes();

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void NearlyTransparentOutlineDoesNotProveAField() {
        byte[] source = StaticPdf("q /GS1 gs 0 0 0 RG 1 w 80 80 120 20 re S Q",
            "/Resources << /ExtGState << /GS1 << /CA 0.001 >> >> >>");
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label).Proposals);
    }

    [Fact]
    public void TightClipAroundUnderlineStillAllowsItsInferredField() {
        byte[] source = StaticPdf("q 79 118 122 3 re W n 0 0 0 RG 1 w 80 120 m 200 120 l S Q");
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 60D, 70D, 80D, 1D) };

        PdfStaticFormFieldProposal proposal = Assert.Single(
            PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label).Proposals);
        Assert.Equal(PdfStaticFormEvidenceKind.Underline, proposal.EvidenceKind);
    }

    [Fact]
    public void BatchedRectanglesProvideIndependentFieldEvidence() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re 80 40 120 20 re S");
        var labels = new[] {
            new PdfStaticFormTextEvidence(1, "First", 10D, 100D, 70D, 120D, 1D),
            new PdfStaticFormTextEvidence(1, "Second", 10D, 140D, 70D, 160D, 1D)
        };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Equal(2, report.Proposals.Count);
    }

    [Fact]
    public void BatchedHorizontalRulesProvideIndependentFieldEvidence() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 120 m 200 120 l 80 80 m 200 80 l S");
        var labels = new[] {
            new PdfStaticFormTextEvidence(1, "First", 10D, 60D, 70D, 80D, 1D),
            new PdfStaticFormTextEvidence(1, "Second", 10D, 100D, 70D, 120D, 1D)
        };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Equal(2, report.Proposals.Count);
    }

    [Fact]
    public void ReadingOrderGroupsSlightlyOffsetRightToLeftFieldsIntoOneRow() {
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Shape(Box(80D, 20D), 70D, 82D)
                .Shape(Box(80D, 20D), 250D, 80D)).ToBytes();
        var labels = new[] {
            new PdfStaticFormTextEvidence(1, "\u05E9\u05DD", 70D, 45D, 140D, 65D, 1D),
            new PdfStaticFormTextEvidence(1, "\u05E2\u05D9\u05E8", 250D, 45D, 320D, 65D, 1D)
        };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Equal(2, report.Proposals.Count);
        Assert.True(report.Proposals[0].VisualBounds.Left > report.Proposals[1].VisualBounds.Left);
    }

    [Fact]
    public void NonrectangularImageClipOutsideFieldDoesNotOccupyIt() {
        const string content = "0 0 0 RG 1 w 80 80 120 20 re S q 190 101 m 210 101 l 210 120 l h W n 150 0 0 60 70 70 cm /Im1 Do Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>",
            "stream", "AAA", "endstream", "endobj", "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        Assert.Single(PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label).Proposals);
    }

    [Fact]
    public void LaterWhiteFillCoveringOnlyTheLargeMarkRestoresAnEmptyField() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S 0 0 0 rg 82 81 116 18 re f 1 1 1 rg 82 81 116 18 re f");
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);
        Assert.True(report.Proposals.Count == 1,
            "Expected an erased mark; diagnostics: " + string.Join(", ", report.Diagnostics.Select(static diagnostic => diagnostic.Code)));
    }

    [Fact]
    public void LargePaintedValueReportsAnOccupiedField() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S 0 0 0 rg 82 81 116 18 re f");
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Theory]
    [InlineData(-1D, 0D, 10D, 10D)]
    [InlineData(0D, -1D, 10D, 10D)]
    public void OcrEvidenceRejectsNegativeVisualOrigin(double left, double top, double right, double bottom) {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfStaticFormTextEvidence(1, "Name", left, top, right, bottom, 1D));
    }

    private static byte[] StaticPdf(string content, string resources = "") =>
        System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R " + resources + "/Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));

    [Fact]
    public void DarkOutlineOnMatchingOpaqueBackdropIsNotVisible() {
        OfficeShape backdrop = OfficeShape.Rectangle(400D, 300D);
        backdrop.FillColor = OfficeColor.Black;
        backdrop.StrokeColor = null;
        OfficeShape outline = Box(140D, 20D);
        outline.FillColor = null;
        outline.StrokeColor = OfficeColor.Black;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Shape(backdrop, 0D, 0D).Shape(outline, 100D, 28D)).ToBytes();
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 20D, 28D, 70D, 48D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "invisible-outline");
    }

    [Fact]
    public void OpaqueWhiteFieldFillMakesItsDarkOutlineVisibleOnDarkPage() {
        const string content = "0 0 0 rg 0 0 240 200 re f 1 1 1 rg 0 0 0 RG 1 w 80 80 120 20 re B";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);
        Assert.True(report.Proposals.Count == 1,
            "Expected a visible empty field; diagnostics: " + string.Join(", ", report.Diagnostics.Select(static diagnostic => diagnostic.Code)));
    }

    [Fact]
    public void SameSizeDarkBackingCanBeAnEmptyOutlinedField() {
        OfficeShape backing = OfficeShape.Rectangle(140D, 20D);
        backing.FillColor = OfficeColor.Black;
        backing.StrokeColor = null;
        OfficeShape outline = Box(140D, 20D);
        outline.FillColor = null;
        outline.StrokeColor = OfficeColor.White;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Shape(backing, 100D, 28D).Shape(outline, 100D, 28D)).ToBytes();
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 20D, 28D, 70D, 48D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);
        Assert.True(report.Proposals.Count == 1,
            "Expected a visible empty field; diagnostics: " + string.Join(", ", report.Diagnostics.Select(static diagnostic => diagnostic.Code)));
    }

    [Fact]
    public void OpaqueWhiteFillCanEraseTheContrastOfAWhiteOutlineOnDarkBacking() {
        const string content = "0 0 0 rg 80 80 120 20 re f 1 1 1 rg 1 1 1 RG 1 w 80 80 120 20 re B";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "invisible-outline");
    }

    [Fact]
    public void DifferenceBlendedWhiteInteriorMarkOccupiesStaticCheckbox() {
        const string content = "0 0 0 RG 1 w 100 100 15 15 re S q /GS1 gs 1 1 1 rg 106 105 3 5 re f Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Resources << /ExtGState << /GS1 << /BM /Difference >> >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Agree", 130D, 85D, 200D, 105D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void ScaledStrokeCrossingCheckboxInteriorIsAnOccupiedMark() {
        const string content = "0 0 0 RG 1 w 100 100 15 15 re S q 5 0 0 5 0 0 cm 0.4 w 19.9 21.5 m 19.9 21.8 l S Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Agree", 130D, 85D, 200D, 105D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void NonuniformlyScaledStrokeUsesItsMaximumVisibleEnvelope() {
        const string content = "0 0 0 RG 1 w 100 100 15 15 re S q 10 0 0 0.1 0 0 cm 0.4 w 9.85 1050 m 9.85 1100 l S Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Agree", 130D, 85D, 200D, 105D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void ThickRectangleStrokeOutsideCheckboxGeometryOccupiesItsInterior() {
        const string content = "0 0 0 RG 1 w 100 100 15 15 re S q 4 w 97 105 2 5 re S Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Agree", 130D, 85D, 200D, 105D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void OpaqueImageOverNativeLabelDoesNotSupportAFieldProposal() {
        byte[] image = PdfPngTestImages.CreateRgbPng(80, 30);
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(Box(140D, 20D), 100D, 28D)
                .Image(image, 10D, 20D, 95D, 35D)).ToBytes();

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void JpxEmbeddedAlphaIsNotAnOpaqueNativeLabelCover() {
        const string content = "BT /F1 12 Tf 20 160 Td (Name:) Tj ET 1 w 100 145 120 20 re S q 90 0 0 30 10 145 cm /Im1 Do Q";
        const string jpx = "unsupported";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 6 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /JPXDecode /SMaskInData 1 /Length " + jpx.Length + " >>",
            "stream", jpx, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF", ""
        }));

        PdfExtractedImage extracted = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetImages());
        Assert.True(extracted.HasTransparencyMask);
        Assert.Single(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

}
