using System;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(true, "Content")]
    [InlineData(false, "Separator")]
    [InlineData(true, "Continuation")]
    [InlineData(false, "Notice")]
    public void DefinitionRejectsEffectiveAnchoredParagraphContentWithoutAppendingCapturedBlocks(bool inherited, string location) {
        PdfOptions options = Options();
        if (inherited) options.DefaultParagraphStyle = AnchoredNoteParagraphStyle();
        var document = PdfDocument.Create(options);
        document.Content.Text("ExistingBody", style: Style());
        Action<PdfContentBuilder> unsupported = body => body.Text("RejectedConstraint", style: inherited ? null : AnchoredNoteParagraphStyle());
        Action<PdfContentBuilder> valid = body => body.Text("NoteBody", style: Style());
        var noteOptions = new PdfFootnoteOptions {
            Separator = location == "Separator" ? unsupported : null,
            ContinuationSeparator = location == "Continuation" ? unsupported : null,
            ContinuationNotice = location == "Notice" ? unsupported : null
        };

        Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(location == "Content" ? unsupported : valid, noteOptions));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal("ExistingBody", pdf.GetPage(1).Text);
    }

    [Fact]
    public void ReferencingPageDefaultsCannotIntroduceAnchoredNoteParagraphContent() {
        var document = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Text("RejectedConstraint"));
        document.Compose(pages => pages.Page(page => {
            page.DefaultParagraphStyle(AnchoredNoteParagraphStyle());
            page.Content(body => body.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("Reference").WithFootnoteReference(note)
            }), style: Style()));
        }));

        Assert.Throws<NotSupportedException>(() => document.ToBytes());
    }

    [Fact]
    public void ExplicitNoteStylesPreserveTheirOwnConstraintsWhenDefaultsKeepWithNext() {
        PdfOptions options = Options();
        options.DefaultParagraphStyle = new PdfParagraphStyle { KeepWithNext = true };
        options.DefaultImageStyle = new PdfImageStyle { KeepWithNext = true };
        options.DefaultHorizontalRuleStyle = new PdfHorizontalRuleStyle { KeepWithNext = true };
        var document = PdfDocument.Create(options);
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            body.Text("SupportedNote", style: Style());
            body.Image(PdfPngTestImages.Create16BitRgbPng(), 4, 4, style: new PdfImageStyle { SpacingBefore = 0, SpacingAfter = 0 });
            body.HR(style: new PdfHorizontalRuleStyle { SpacingBefore = 0, SpacingAfter = 0 });
        });
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("SupportedNote", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Single(pdf.GetPage(1).GetImages());
    }

    private static PdfParagraphStyle AnchoredNoteParagraphStyle() => new() {
        AnchoredCanvas = new PdfCanvasBlock(new PdfPageCanvas().ForegroundImage(PdfPngTestImages.Create16BitRgbPng(), 0, 0, 4, 4).Items)
    };
}
