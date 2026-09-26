using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfStaticFormRecognizerTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void PartiallyErasedNativeLabelIsNotCompleteEvidence(int coverKind) {
        const string label = "BT /F1 12 Tf 10 85 Td (Name) Tj ET ";
        string cover = coverKind == 2 ? "1 1 1 RG 10 w 35 80 m 35 100 l S" : coverKind == 1 ? "q 20 0 0 20 30 80 cm /Im1 Do Q" : "1 g 30 80 20 20 re f";
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 120 20 re S " + label + cover,
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> /XObject << /Im1 5 0 R >> >> ",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 7 >>\nstream\nFFFFFF>\nendstream\nendobj");

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OneLabelNamesOnlyTheNearestFieldRegardlessOfPaintOrder(bool reverseOrder) {
        const string near = "80 80 60 20 re S ";
        const string far = "145 80 60 20 re S ";
        byte[] source = StaticPdf("0 0 0 RG 1 w " + (reverseOrder ? far + near : near + far));

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 105D, 60D, 120D, 1D) });

        PdfStaticFormFieldProposal proposal = Assert.Single(report.Proposals);
        Assert.Equal(80D, proposal.VisualBounds.Left);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DetectedTableContentDoesNotSupplyFormEvidence(bool outsideTable) {
        string content = "0 0 0 RG 1 w " + (outsideTable ? "160 120 60 20 re S " : "80 60 120 20 re S ") +
            "BT /F1 12 Tf 15 125 Td (Label) Tj 70 0 Td (Value) Tj ET " +
            "BT /F1 12 Tf 15 105 Td (Name) Tj 70 0 Td (Alice) Tj ET " +
            "BT /F1 12 Tf 15 85 Td (State) Tj 70 0 Td (WA) Tj ET " +
            "BT /F1 12 Tf 15 65 Td (City) Tj ET";
        byte[] source = StaticPdf(content,
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        PdfDocument document = PdfDocument.Load(source);
        PdfDocumentReadResult read = document.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast });
        Assert.NotEmpty(read.Pages[0].Tables);
        Assert.Contains(read.Pages[0].TextBlocks, static block => block.IsTableContent);

        var ocr = outsideTable ? null : new[] { new PdfStaticFormTextEvidence(1, "City", 15D, 125D, 50D, 140D, 1D) };
        Assert.Empty(document.Forms.RecognizeStaticLayout(ocrText: ocr).Proposals);
    }
    [Fact]
    public void EquallySharedAboveLabelRequiresManualAssignment() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 60 20 re S 145 80 60 20 re S");
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] { new PdfStaticFormTextEvidence(1, "Name", 80D, 80D, 205D, 95D, 1D) });
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "ambiguous-label");
    }

    [Fact]
    public void DistinctAboveLabelsRemainIndependent() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 60 20 re S 145 80 60 20 re S");
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] {
                new PdfStaticFormTextEvidence(1, "First name", 80D, 80D, 140D, 95D, 1D),
                new PdfStaticFormTextEvidence(1, "Last name", 145D, 80D, 205D, 95D, 1D)
            });
        Assert.Equal(2, report.Proposals.Count);
        Assert.Equal(new[] { "First name", "Last name" }, report.Proposals.Select(static proposal => proposal.Label));
    }

    [Fact]
    public void FullyErasedSpanCannotLeaveAnIncompleteNativeLabel() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 100 80 100 20 re S " +
            "BT /F1 12 Tf 10 85 Td (Last ) Tj 40 0 Td (name) Tj ET q 40 0 0 40 5 65 cm /Im1 Do Q",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> /XObject << /Im1 5 0 R >> >> ",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 7 >>\nstream\nFFFFFF>\nendstream\nendobj");
        PdfDocument document = PdfDocument.Load(source);
        PdfDocumentReadResult read = document.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast });
        Assert.Contains(read.Pages[0].TextBlocks, static block => block.Spans.Count > 1);
        Assert.Empty(document.Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void OverlappingOcrDetectionsCannotNameTwoFields() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 60 20 re S 145 80 60 20 re S");
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] {
                new PdfStaticFormTextEvidence(1, "Name", 130D, 80D, 150D, 95D, 1D),
                new PdfStaticFormTextEvidence(1, "Name", 135D, 80D, 155D, 95D, 1D)
            });
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "ambiguous-label");
    }

}
