using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfStaticFormRecognizerTests {
    [Fact]
    public void ProposalLimitCountsFieldsRetainedAfterLabelAssignment() {
        byte[] pdf = StaticPdf("0 G 1 w 80 80 120 20 re S 80 50 120 20 re S");
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };
        var report = PdfDocument.Load(pdf).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MaxProposals = 1 }, labels);
        Assert.Equal(100D, Assert.Single(report.Proposals).VisualBounds.Top);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ApplyingRecognizedFieldsPreservesArtifactReadPolicy(bool includeArtifacts) {
        byte[] pdf = StaticPdf("0 G 1 w 80 80 120 20 re S /Artifact BMC BT /F1 12 Tf 10 160 Td (Footer) Tj ET EMC",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        PdfDocument source = PdfDocument.Load(pdf, new PdfLoadOptions { IncludeArtifactText = includeArtifacts });
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };
        var report = source.Forms.RecognizeStaticLayout(ocrText: labels);
        PdfDocument edited = report.ApplySelected(new[] { Assert.Single(report.Proposals).Index }).ToDocument();
        Assert.Equal(includeArtifacts, edited.ReadOptions.IncludeArtifactText);
        Assert.Equal(includeArtifacts, edited.Read().Pages.SelectMany(static page => page.TextBlocks)
            .Any(static block => block.Text.Contains("Footer")));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void DashedStrokeLabelsNeedIndependentEvidence(bool inherited, bool reset) {
        string text = (reset ? "[] 0 d " : "") + "BT /F1 12 Tf 1 Tr 10 85 Td (Name) Tj ET";
        const string fonts = "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >>";
        string objects = "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + fonts +
            " >> /Length " + text.Length + " >>\nstream\n" + text + "\nendstream\nendobj";
        byte[] pdf = StaticPdf("0 G 1 w 80 80 120 20 re S q [1 1000] 100 d " + (inherited ? "/Fm1 Do" : text) + " Q",
            "/Resources << " + fonts + " /XObject << /Fm1 5 0 R >> >> ", objects);
        Assert.Equal(reset ? 1 : 0, PdfDocument.Load(pdf).Forms.RecognizeStaticLayout().Proposals.Count);
    }

    [Fact]
    public void FilledAreaComparisonsCannotBypassTheCandidateScanBudget() {
        string backing = string.Concat(Enumerable.Repeat("1 g 0 0 240 200 re f ", 12));
        byte[] pdf = StaticPdf(backing + "0 G 1 w 80 80 120 20 re S");
        var labels = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };
        var error = Assert.Throws<PdfReadLimitException>(() => PdfDocument.Load(pdf).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MaxCandidateScanWork = 1000 }, labels));
        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, error.Kind);
        Assert.Single(PdfDocument.Load(pdf).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MaxCandidateScanWork = 100000 }, labels).Proposals);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FormLabelsRetainInheritedTintResolution(bool nested, bool resolved) {
        const string text = "BT /F1 12 Tf 1 Tr 10 85 Td (Name) Tj ET";
        const string fonts = "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >>";
        string form = nested ? "/Child Do" : text;
        string resources = nested ? "/XObject << /Child 7 0 R >>" : fonts;
        string objects = "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + resources +
            " >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj\n";
        const string tint = "{ 1 exch div 1 exch sub }";
        objects += resolved
            ? "6 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>\nendobj\n"
            : "6 0 obj\n<< /FunctionType 4 /Domain [0 1] /Range [0 1] /Length " + tint.Length +
                " >>\nstream\n" + tint + "\nendstream\nendobj\n";
        if (nested) objects += "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + fonts +
            " >> /Length " + text.Length + " >>\nstream\n" + text + "\nendstream\nendobj";
        byte[] pdf = StaticPdf("0 G 1 w 80 80 120 20 re S /Tint CS " + (resolved ? "1" : "0") + " SCN /Fm1 Do",
            "/Resources << /XObject << /Fm1 5 0 R >> /ColorSpace << /Tint [/Separation /Brand /DeviceGray 6 0 R] >> >> ", objects);
        var report = PdfDocument.Load(pdf).Forms.RecognizeStaticLayout();
        Assert.Equal(resolved ? 1 : 0, report.Proposals.Count);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    public void ColorSpaceSelectionResetsLabelPaint(int placement, bool stroke) {
        string text = "/Tint " + (stroke ? "CS" : "cs") + " BT /F1 12 Tf " +
            (stroke ? "1" : "0") + " Tr 10 85 Td (Name) Tj ET";
        const string resources = "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> " +
            "/ColorSpace << /Tint [/Separation /Brand /DeviceGray << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [1] /N 1 >>] >>";
        string form = placement == 2 ? "/Child Do" : text;
        string objects = "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + resources +
            " /XObject << /Child 6 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + resources +
            " >> /Length " + text.Length + " >>\nstream\n" + text + "\nendstream\nendobj";
        byte[] pdf = StaticPdf("0 g 0 G 1 w 80 80 120 20 re S " + (placement == 0 ? text : "/Fm1 Do"),
            "/Resources << " + resources + " /XObject << /Fm1 5 0 R >> >> ", objects);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.White, span.Color);
        Assert.Empty(PdfDocument.Load(pdf).Forms.RecognizeStaticLayout().Proposals);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void FailedColorSpaceLookupCannotReviveThePreviousPaint(bool stroke, bool assignColor) {
        string channel = stroke ? "CS" : "cs";
        string select = stroke ? "SCN" : "scn";
        byte[] pdf = StaticPdf("0 G 1 w 80 80 120 20 re S /Tint " + channel + " 1 " + select +
            " /Missing " + channel + (assignColor ? " 0 " + select : " /RelativeColorimetric ri") + " BT /F1 12 Tf " + (stroke ? "1" : "0") + " Tr 10 85 Td (Name) Tj ET",
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> " +
            "/ColorSpace << /Tint [/Separation /Brand /DeviceGray << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>] >> >> ");
        Assert.Empty(PdfDocument.Load(pdf).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void NativeTextContrastConsumesBudgetWithoutFieldCandidates() {
        string fills = string.Concat(Enumerable.Repeat("1 g 0 0 240 200 re f ", 100));
        string text = string.Concat(Enumerable.Range(0, 4).Select(index =>
            "0 g BT /F1 12 Tf 10 " + (160 - index * 30) + " Td (Name) Tj ET "));
        byte[] pdf = StaticPdf(fills + text,
            "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> ");
        var error = Assert.Throws<PdfReadLimitException>(() => PdfDocument.Load(pdf).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MaxCandidateScanWork = 1500 }));
        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, error.Kind);
        Assert.Empty(PdfDocument.Load(pdf).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MaxCandidateScanWork = 10000 }).Proposals);
    }
}
