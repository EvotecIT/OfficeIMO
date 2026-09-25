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

        PdfAcroFormEditResult edited = report.ApplySelected(source, report.Proposals.Select(static proposal => proposal.Index).ToArray());
        PdfFormField[] fields = PdfInspector.Inspect(edited.ToBytes()).FormFields.ToArray();
        Assert.Equal(2, fields.Length);
        Assert.All(report.Proposals, proposal => Assert.Contains(fields, field => field.Name == proposal.SuggestedName));
        Assert.All(fields, static field => Assert.Single(field.Widgets));
    }

    [Fact]
    public void OcrLabelCanSupportAProposalButCannotAuthorizeAChangedArtifact() {
        byte[] sourceBytes = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(Box(140D, 20D), 100D, 30D))
            .ToBytes();
        var ocr = new[] { new PdfStaticFormTextEvidence(1, "Customer ID", 20D, 30D, 85D, 49D, 0.95D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(sourceBytes).Forms.RecognizeStaticLayout(ocrText: ocr);

        PdfStaticFormFieldProposal proposal = Assert.Single(report.Proposals);
        Assert.True(proposal.UsedOcrLabel);
        Assert.Equal("Customer ID", proposal.Label);

        byte[] changed = PdfDocument.Create().Canvas(canvas => canvas.Text("Different source", 20D, 20D, 120D, 20D)).ToBytes();
        Assert.Throws<InvalidOperationException>(() => report.ApplySelected(PdfDocument.Load(changed), new[] { proposal.Index }));
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
