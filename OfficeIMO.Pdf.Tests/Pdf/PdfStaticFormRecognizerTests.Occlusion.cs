using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfStaticFormRecognizerTests {
    [Theory]
    [InlineData("79 79 m 201 79 l 201 101 l 79 101 l h f")]
    [InlineData("79 79 m 201 79 l 201 101 l 79 101 l f")]
    public void OpaqueRectangularPathErasesAnEarlierOutline(string cover) {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S 1 g " + cover);

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occluded-outline");
    }

    [Theory]
    [InlineData("79 79 2 22 re f", false)]
    [InlineData("199 79 2 22 re f", false)]
    [InlineData("79 79 122 2 re f", false)]
    [InlineData("79 99 122 2 re f", false)]
    [InlineData("82 82 116 16 re f", true)]
    public void LaterFillMustLeaveEveryOutlineSideVisible(string cover, bool intact) {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S 1 g " + cover);

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        Assert.Equal(intact ? 1 : 0, report.Proposals.Count);
        if (!intact) Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occluded-outline");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PartialWritingLineRepaintRespectsPaintOrder(bool paintedBefore) {
        const string outline = "0 0 0 RG 1 w 80 80 m 200 80 l S ";
        const string cover = "1 g 110 79 20 2 re f ";
        byte[] source = StaticPdf(paintedBefore ? cover + outline : outline + cover);

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        Assert.Equal(paintedBefore ? 1 : 0, report.Proposals.Count);
        if (!paintedBefore) Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occluded-outline");
    }

    [Theory]
    [InlineData("60 60 160 60 re 80 80 120 20 re B*")]
    [InlineData("60 60 160 60 re 80 80 m 80 100 l 200 100 l 200 80 l h B")]
    public void CompoundFillHolesCannotEraseAnEarlierFieldValue(string compound) {
        byte[] source = StaticPdf("0 g 110 87 5 5 re f 1 g 0 0 0 RG 1 w " + compound);

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        Assert.Empty(report.Proposals);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ThinImageCoverMustLeaveTheOutlineVisible(bool paintedBefore) {
        const string outline = "0 0 0 RG 0.2 w 80 80 120 20 re S ";
        const string cover = "q 0.4 0 0 22 79.8 79 cm /Im1 Do Q ";
        byte[] source = StaticPdf(paintedBefore ? cover + outline : outline + cover,
            "/Resources << /XObject << /Im1 5 0 R >> >> ",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 7 >>\nstream\nFFFFFF>\nendstream\nendobj");

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        Assert.Equal(paintedBefore ? 1 : 0, report.Proposals.Count);
        if (!paintedBefore) Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occluded-outline");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ThinStrokeCoverMustLeaveTheOutlineVisible(bool paintedBefore) {
        const string outline = "0 0 0 RG 0.2 w 80 80 120 20 re S ";
        const string cover = "1 1 1 RG 0.2 w 80 87 m 80 92 l S ";
        byte[] source = StaticPdf(paintedBefore ? cover + outline : outline + cover);

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        Assert.Equal(paintedBefore ? 1 : 0, report.Proposals.Count);
        if (!paintedBefore) Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occluded-outline");
    }
    [Fact]
    public void LaterSurroundingFrameDoesNotEraseTheInnerField() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S 70 70 140 40 re S");

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        Assert.Single(report.Proposals);
    }

}
