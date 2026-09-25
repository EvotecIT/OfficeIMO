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
