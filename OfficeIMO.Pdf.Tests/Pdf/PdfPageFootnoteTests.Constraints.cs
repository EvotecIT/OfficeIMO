using System;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData("ParagraphKeep", true, "Content")]
    [InlineData("ParagraphCanvas", true, "Content")]
    [InlineData("Image", false, "Content")]
    [InlineData("Image", true, "Separator")]
    [InlineData("Rule", false, "Continuation")]
    [InlineData("Rule", true, "Content")]
    public void DefinitionRejectsUnsupportedEffectiveConstraintsWithoutAppendingCapturedContent(string constraint, bool inherited, string location) {
        PdfOptions options = Options();
        if (inherited) ApplyUnsupportedNoteDefault(options, constraint);
        var document = PdfDocument.Create(options);
        document.Content.Text("ExistingBody", style: Style());
        Action<PdfContentBuilder> unsupported = body => AddConstrainedNoteContent(body, constraint, inherited);
        Action<PdfContentBuilder> valid = body => body.Text("NoteBody", style: Style());
        var noteOptions = new PdfFootnoteOptions {
            Separator = location == "Separator" ? unsupported : null,
            ContinuationSeparator = location == "Continuation" ? unsupported : null
        };

        Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(location == "Content" ? unsupported : valid, noteOptions));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal("ExistingBody", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData("ParagraphKeep")]
    [InlineData("ParagraphCanvas")]
    [InlineData("Image")]
    [InlineData("Rule")]
    public void ReferencingPageDefaultsCannotIntroduceUnsupportedNoteConstraints(string constraint) {
        var document = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => AddConstrainedNoteContent(body, constraint, inherited: true));
        PdfOptions defaults = Options();
        ApplyUnsupportedNoteDefault(defaults, constraint);
        document.Compose(pages => pages.Page(page => {
            if (constraint.StartsWith("Paragraph", StringComparison.Ordinal)) page.DefaultParagraphStyle(defaults.DefaultParagraphStyle!);
            else if (constraint == "Image") page.DefaultImageStyle(defaults.DefaultImageStyle!);
            else page.DefaultHorizontalRuleStyle(defaults.DefaultHorizontalRuleStyle!);
            page.Content(body => body.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("Reference").WithFootnoteReference(note)
            }), style: Style()));
        }));

        Assert.Throws<NotSupportedException>(() => document.ToBytes());
    }

    [Fact]
    public void ExplicitSupportedNoteStylesOverrideUnsupportedDefaultKeeps() {
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

    private static void AddConstrainedNoteContent(PdfContentBuilder body, string constraint, bool inherited) {
        if (constraint.StartsWith("Paragraph", StringComparison.Ordinal)) {
            body.Text("RejectedConstraint", style: inherited ? null : new PdfParagraphStyle { KeepWithNext = true });
        } else if (constraint == "Image") {
            body.Image(PdfPngTestImages.Create16BitRgbPng(), 4, 4, style: inherited ? null : new PdfImageStyle { KeepWithNext = true });
            body.Text("RejectedConstraint", style: Style());
        } else {
            body.HR(style: inherited ? null : new PdfHorizontalRuleStyle { KeepWithNext = true });
            body.Text("RejectedConstraint", style: Style());
        }
    }

    private static void ApplyUnsupportedNoteDefault(PdfOptions options, string constraint) {
        if (constraint == "ParagraphKeep") options.DefaultParagraphStyle = new PdfParagraphStyle { KeepWithNext = true };
        else if (constraint == "ParagraphCanvas") options.DefaultParagraphStyle = new PdfParagraphStyle {
            AnchoredCanvas = new PdfCanvasBlock(new PdfPageCanvas().ForegroundImage(PdfPngTestImages.Create16BitRgbPng(), 0, 0, 4, 4).Items)
        };
        else if (constraint == "Image") options.DefaultImageStyle = new PdfImageStyle { KeepWithNext = true };
        else options.DefaultHorizontalRuleStyle = new PdfHorizontalRuleStyle { KeepWithNext = true };
    }
}
