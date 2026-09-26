using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfStaticFormRecognizerTests {
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
