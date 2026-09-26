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

    [Theory]
    [InlineData("Name")]
    [InlineData("Nane")]
    public void OverlappingOcrDetectionsCannotNameTwoFields(string secondText) {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 80 60 20 re S 145 80 60 20 re S");
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            ocrText: new[] {
                new PdfStaticFormTextEvidence(1, "Name", 130D, 80D, 150D, 95D, 1D),
                new PdfStaticFormTextEvidence(1, secondText, 135D, 80D, 155D, 95D, 1D)
            });
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "ambiguous-label");
    }

    [Theory]
    [InlineData(0.5D)]
    [InlineData(1D)]
    public void DuplicateOcrPreservesStrongerNativeConfidenceAndProvenance(double ocrConfidence) {
        byte[] source = StaticPdf("0 0 0 RG 1 w 80 120 m 200 120 l S BT /F1 12 Tf 10 125 Td (Name) Tj ET",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        PdfDocument document = PdfDocument.Load(source);
        PdfStaticFormFieldProposal original = Assert.Single(document.Forms.RecognizeStaticLayout().Proposals);
        var options = new PdfStaticFormRecognitionOptions { MinimumConfidence = original.Confidence - 0.01D };
        var duplicate = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 69D, 43D, 87D, ocrConfidence) };
        PdfStaticFormFieldProposal repeated = Assert.Single(document.Forms.RecognizeStaticLayout(options, duplicate).Proposals);
        Assert.Equal(original.Confidence, repeated.Confidence, 10);
        Assert.False(repeated.UsedOcrLabel);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactPartialTextClipRetainsOnlyVisibleOccupancy(bool visibleInsideField) {
        int clipX = visibleInsideField ? 120 : 60;
        byte[] source = StaticPdf("0 0 0 RG 1 w 100 80 100 20 re S q " + clipX +
            " 60 15 60 re W n BT /F1 12 Tf 60 85 Td (Value extending into field) Tj ET Q",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 45D, 120D, 1D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Equal(visibleInsideField ? 0 : 1, report.Proposals.Count);
        Assert.Equal(visibleInsideField, report.Diagnostics.Any(static diagnostic => diagnostic.Code == "occupied-field"));
    }

    [Fact]
    public void FullyRepaintedPartialTextClipDoesNotOccupyTheField() {
        byte[] source = StaticPdf("0 0 0 RG 1 w 100 80 100 20 re S q 120 81 15 18 re W n " +
            "BT /F1 12 Tf 60 85 Td (Value extending into field) Tj ET Q 1 g 118 81 19 18 re f",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 45D, 120D, 1D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Single(report.Proposals);
        Assert.DoesNotContain(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(2, 1)]
    [InlineData(1, 30)]
    [InlineData(5, 30)]
    [InlineData(2, 30)]
    [InlineData(6, 30)]
    public void TextStrokeEnvelopeOccupiesTheField(int renderingMode, int strokeWidth) {
        byte[] source = StaticPdf("0 0 0 RG 1 w 100 80 100 20 re S BT /F1 12 Tf " +
            renderingMode + " Tr " + strokeWidth + " w 60 85 Td (Name) Tj ET",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 45D, 120D, 1D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Equal(strokeWidth >= 30 ? 0 : 1, report.Proposals.Count);
        Assert.Equal(strokeWidth >= 30, report.Diagnostics.Any(static diagnostic => diagnostic.Code == "occupied-field"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void FormTextStrokeEnvelopeUsesInheritedAndLocalState(int sourceKind) {
        string text = (sourceKind == 3 ? "1 w " : "") + "BT /F1 12 Tf 1 Tr 60 85 Td (Name) Tj ET";
        const string fonts = "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >>";
        string form = sourceKind == 1 ? "/Child Do" : text;
        string formResources = sourceKind == 1 ? "/XObject << /Child 6 0 R >>" : fonts;
        string objects = "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + formResources +
            " >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj";
        if (sourceKind == 1) objects += "\n6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + fonts +
            " >> /Length " + text.Length + " >>\nstream\n" + text + "\nendstream\nendobj";
        byte[] source = StaticPdf("0 0 0 RG 1 w 100 80 100 20 re S " +
            (sourceKind == 2 ? "/GS1 gs " : sourceKind >= 4 ? "1 w 30 M " + (sourceKind == 5 ? "1 j " : "") : "30 w q 1 w Q ") + "/Fm1 Do",
            "/Resources << /XObject << /Fm1 5 0 R >> /ExtGState << /GS1 << /Type /ExtGState /LW 30 >> >> >> ", objects);
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 45D, 120D, 1D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);
        Assert.Equal(sourceKind == 3 || sourceKind == 5 ? 1 : 0, report.Proposals.Count);
        Assert.Equal(sourceKind != 3 && sourceKind != 5, report.Diagnostics.Any(static diagnostic => diagnostic.Code == "occupied-field"));
    }

}
