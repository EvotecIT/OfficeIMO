using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStaticFormRecognizerTests {
    [Fact]
    public void ProposalsRequireSelectionAndCreateReopenableFields() {
        byte[] sourceBytes = CreateStaticForm();
        PdfDocument source = PdfDocument.Load(sourceBytes);

        PdfStaticFormRecognitionReport report = source.Forms.RecognizeStaticLayout();

        Assert.Equal(2, report.Proposals.Count);
        Assert.Equal(new[] { PdfFormFieldCreationKind.Text, PdfFormFieldCreationKind.CheckBox },
            report.Proposals.Select(static proposal => proposal.Kind));
        Assert.Empty(PdfInspector.Inspect(sourceBytes).FormFields);

        PdfAcroFormEditResult edited = report.ApplySelected(report.Proposals.Select(static proposal => proposal.Index).ToArray());
        PdfFormField[] fields = PdfInspector.Inspect(edited.ToBytes()).FormFields.ToArray();
        Assert.Equal(2, fields.Length);
        Assert.All(report.Proposals, proposal => Assert.Contains(fields, field => field.Name == proposal.SuggestedName));
        Assert.All(fields, static field => Assert.Single(field.Widgets));
        PdfVisualComparisonReport visual = source.Proof.CompareVisual(edited.ToDocument());
        OfficeRasterImage rendered = VisualBaselineTestSupport.DecodePng(visual.Pages[0].ActualPng, "Edited form page must render as PNG.");
        Assert.True(rendered.GetPixel(110, 27).R < 200, "The static text-field outline must remain visible below its widget.");
        Assert.True(rendered.GetPixel(110, 35).R > 245, "The proposed widget must leave the writing area unobscured.");
    }

    [Fact]
    public void OcrLabelCanSupportAProposalFromAnExactSourceSnapshot() {
        byte[] sourceBytes = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(Box(140D, 20D), 100D, 30D))
            .ToBytes();
        var ocr = new[] { new PdfStaticFormTextEvidence(1, "Customer ID", 20D, 30D, 85D, 49D, 0.95D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(sourceBytes).Forms.RecognizeStaticLayout(ocrText: ocr);

        PdfStaticFormFieldProposal proposal = Assert.Single(report.Proposals);
        Assert.True(proposal.UsedOcrLabel);
        Assert.Equal("Customer ID", proposal.Label);

        Assert.Equal(PdfArtifactFingerprint.ComputeSha256(sourceBytes), report.SourceSha256);
        Assert.Contains(report.ApplySelected(new[] { proposal.Index }).Fields, field => field.Name == proposal.SuggestedName);
    }

    [Fact]
    public void SuggestedNamesReserveFieldAncestorsAcrossUnselectedPages() {
        byte[] sourceBytes = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Text("Customer:", 20D, 28D, 85D, 20D).Shape(Box(140D, 20D), 125D, 28D))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Existing field lives here"))
            .ToBytes();
        PdfAcroFormEditResult withExistingField = PdfDocument.Load(sourceBytes).Forms.Edit(edit => edit.Create(
            new PdfFormFieldCreateOptions { Name = "customer.name", PageNumber = 2, X = 20D, Y = 20D, Width = 100D, Height = 20D }));

        PdfStaticFormRecognitionReport report = withExistingField.ToDocument().Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { PageSelection = PdfPageSelection.From(1) });

        PdfStaticFormFieldProposal proposal = Assert.Single(report.Proposals);
        Assert.Equal("customer_2", proposal.SuggestedName);
        Assert.Contains(report.ApplySelected(new[] { proposal.Index }).Fields, field => field.Name == "customer_2");
    }

    [Fact]
    public void GeneratedEncryptedSourceRetainsSnapshotAndHonorsMutationGate() {
        PdfDocument source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 }
            .SetEncryption(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" }))
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D).Shape(Box(140D, 20D), 100D, 28D));

        PdfStaticFormRecognitionReport report = source.Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
        Assert.Throws<PdfMutationBlockedException>(() => report.ApplySelected(new[] { 0 }));
    }

    [Fact]
    public void ProposedWidgetLeavesSourceMarksVisible() {
        PdfStaticFormRecognitionReport report = PdfDocument.Load(CreateStaticForm()).Forms.RecognizeStaticLayout();

        Assert.All(report.Proposals, proposal => {
            PdfFormFieldStyle style = proposal.ToCreateOptions().Style!;
            Assert.Null(style.BackgroundColor);
            Assert.Null(style.BorderColor);
            Assert.Equal(0D, style.BorderWidth);
        });
    }

    [Fact]
    public void LowConfidenceCandidatesRemainVisibleInDiagnostics() {
        PdfStaticFormRecognitionReport report = PdfDocument.Load(CreateStaticForm()).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MinimumConfidence = 0.8D });

        Assert.DoesNotContain(report.Proposals, static proposal => proposal.Kind == PdfFormFieldCreationKind.CheckBox);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "low-confidence");
    }

    [Fact]
    public void ContainedValueTextAndGradientPaintAreNotEmptyFields() {
        OfficeShape gradientBox = Box(100D, 28D);
        gradientBox.FillColor = null;
        gradientBox.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(Box(100D, 28D), 140D, 20D)
                .Text("Alice", 150D, 25D, 60D, 20D)
                .Text("Status:", 20D, 78D, 70D, 20D)
                .Shape(gradientBox, 140D, 70D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
    }

    [Fact]
    public void CheckedStaticBoxIsNotProposedAsAnEmptyCheckbox() {
        OfficeShape firstMark = OfficeShape.Line(0D, 0D, 9D, 9D);
        firstMark.StrokeColor = OfficeColor.Black;
        OfficeShape secondMark = OfficeShape.Line(0D, 9D, 9D, 0D);
        secondMark.StrokeColor = OfficeColor.Black;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(Box(15D, 15D), 100D, 75D)
                .Shape(firstMark, 103D, 78D)
                .Shape(secondMark, 103D, 78D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void EdgeNearCheckmarkAndSmallFilledDotBothOccupyAStaticBox() {
        OfficeShape mark = OfficeShape.Line(0D, 0D, 13D, 13D);
        mark.StrokeColor = OfficeColor.Black;
        OfficeShape dot = OfficeShape.Rectangle(3D, 3D);
        dot.FillColor = OfficeColor.Black;
        dot.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(Box(15D, 15D), 100D, 75D)
                .Shape(mark, 101D, 76D)
                .Text("First", 125D, 72D, 100D, 20D)
                .Shape(Box(15D, 15D), 100D, 115D)
                .Shape(dot, 106D, 121D)
                .Text("Second", 125D, 112D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Equal(2, report.Diagnostics.Count(static diagnostic => diagnostic.Code == "occupied-field"));
    }

    [Fact]
    public void OpaqueFieldFillCanCoverAnEarlierMark() {
        OfficeShape mark = OfficeShape.Line(0D, 0D, 9D, 9D);
        mark.StrokeColor = OfficeColor.Black;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(mark, 103D, 78D)
                .Shape(Box(15D, 15D), 100D, 75D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
    }

    [Fact]
    public void NonTruncatingRectangularClipKeepsStaticFieldCandidate() {
        const string content = "q 0 0 400 300 re W n 1 w 100 205 15 15 re S BT /F1 12 Tf 125 208 Td (I agree) Tj ET Q\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", string.Empty
        }));

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
    }

    [Theory]
    [InlineData("1 w 103 208 m 112 217 l S")]
    [InlineData("0 0 0 rg 101 206 13 13 re f")]
    public void PaintClippedOutsideCheckboxDoesNotOccupyIt(string clippedPaint) {
        string content = $"q 0 0 10 10 re W n {clippedPaint} Q 1 w 100 205 15 15 re S BT /F1 12 Tf 125 208 Td (I agree) Tj ET\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", string.Empty
        }));

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
    }

    [Fact]
    public void ExistingLinkAnnotationBlocksOverlappingProposal() {
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D).Shape(Box(140D, 20D), 100D, 28D))
            .ToBytes();
        PdfDocument annotated = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Link", LinkUri = "https://example.com", Rectangle = new[] { 100D, 252D, 240D, 272D }
        }).ToDocument();

        PdfStaticFormRecognitionReport report = annotated.Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "existing-annotation");
    }

    [Fact]
    public void DiagnosticLimitReportsTruncation() {
        PdfStaticFormRecognitionReport report = PdfDocument.Load(CreateStaticForm()).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MinimumConfidence = 1D, MaxDiagnostics = 1 });

        Assert.Empty(report.Proposals);
        Assert.Equal("diagnostics-truncated", Assert.Single(report.Diagnostics).Code);
    }

    private static byte[] CreateStaticForm() {
        OfficeShape textBox = Box(140D, 20D);
        OfficeShape checkBox = Box(15D, 15D);
        return PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas
                .Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(textBox, 100D, 28D)
                .Shape(checkBox, 100D, 75D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();
    }

    private static OfficeShape Box(double width, double height) {
        OfficeShape box = OfficeShape.Rectangle(width, height);
        box.FillColor = OfficeColor.White;
        box.StrokeColor = OfficeColor.Black;
        box.StrokeWidth = 1D;
        return box;
    }
}
