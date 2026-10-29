using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void KeepWithNextRetainsOnlyTheTerminalParagraphsFirstLegalFragment(bool inherited) {
        PdfCreateOptions options = Options();
        var kept = NoteLineStyle(20);
        kept.KeepWithNext = true;
        if (inherited) options.DefaultParagraphStyle = kept;
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => body
            .Text("NoteLead", style: NoteLineStyle(20))
            .Text("KeptHeading", style: inherited ? null : kept)
            .Text(string.Join("\n", Enumerable.Range(0, 6).Select(index => $"Terminal{index:D2}")), style: NoteLineStyle(20)));
        document.Content.Spacer(50).Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("NoteLead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("KeptHeading", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("KeptHeading", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("Terminal00", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("Terminal03", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Terminal04", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("Terminal05", pdf.GetPage(3).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true, 0D, 4, 3)]
    [InlineData(false, 0D, 3, 2)]
    [InlineData(false, 4D, 4, 3)]
    public void DecoratedNoteFragmentsReserveTheirActualPaddingAndPreserveTheCapturedStyle(
        bool repeat, double inset, int pages, int markerPage) {
        PdfCreateOptions options = Options();
        options.PageHeight = 200;
        var panelStyle = new PdfPanelStyle {
            Background = global::OfficeIMO.OfficeColor.FromRgb(255, 255, 0), BorderColor = global::OfficeIMO.OfficeColor.Black,
            PaddingX = 6, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0,
            RepeatFragmentDecoration = repeat, FragmentBottomInset = inset
        };
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => body.Panel(panel => panel.Text(
            string.Join("\n", Enumerable.Range(1, 22).Select(index => $"PanelLine{index:D2}")),
            style: NoteLineStyle(20)), panelStyle), new PdfFootnoteOptions { Placement = PdfFootnotePlacement.BeneathText });
        panelStyle.PaddingX = 50;
        panelStyle.PaddingY = 50;
        panelStyle.Background = null;
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(pages, pdf.NumberOfPages);
        Assert.Contains("PanelLine15", pdf.GetPage(markerPage).Text, StringComparison.Ordinal);
        Assert.Equal(Enumerable.Range(1, 22).Select(index => $"PanelLine{index:D2}"), pdf.GetPages()
            .SelectMany(page => page.GetWords()).Where(word => word.Text.StartsWith("PanelLine", StringComparison.Ordinal))
            .Select(word => word.Text));
        var first = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "PanelLine01");
        var continuation = Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "PanelLine08");
        Assert.Equal(26D, first.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(repeat ? 10D : 14D, continuation.BoundingBox.Top - first.BoundingBox.Top, 3);
        foreach (var page in pdf.GetPages()) {
            var fill = Assert.Single(page.Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), bounds => bounds.HasValue);
            Assert.Equal(20D, fill!.Value.Left, 3);
            Assert.Equal(240D, fill.Value.Right, 3);
            Assert.True(fill.Value.Bottom >= 20D - .001D);
            Assert.All(page.GetWords().Where(word => word.Text.StartsWith("PanelLine", StringComparison.Ordinal)), word => {
                Assert.True(word.BoundingBox.Top <= fill.Value.Top + .001D);
                Assert.True(word.BoundingBox.Bottom >= fill.Value.Bottom - .001D);
            });
        }
    }

    [Fact]
    public void NumberedNoteItemsKeepTheirMarkersOnceAcrossContinuationAndNestedFrames() {
        var document = PdfDocument.Create(Options());
        var note = document.Content.DefineFootnote(body => body.Panel(panel => panel.RichNumbered(new[] {
            PdfListItem.Rich(new[] { PdfTextRun.Bolded(string.Join("\n", Enumerable.Range(0, 12).Select(index => $"ListLine{index:D2}"))) }),
            PdfListItem.Plain("SecondItem"), PdfListItem.Plain("ThirdItem")
        }, style: new PdfListStyle {
            FontSize = 8, LineSpacing = PdfLineSpacing.Exactly(10), LeftIndent = 4,
            MarkerWidth = 12, MarkerGap = 3, ItemSpacing = 0, SpacingBefore = 0, SpacingAfter = 0
        }), new PdfPanelStyle { PaddingX = 6, PaddingY = 0, BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0 }));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        foreach (string marker in new[] { "1.", "2.", "3." }) Assert.Single(words, word => word.Text == marker);
        Assert.Contains("1.", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("1.", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("SecondItem", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("ThirdItem", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Equal(Enumerable.Range(0, 12).Select(index => $"ListLine{index:D2}"),
            words.Where(word => word.Text.StartsWith("ListLine", StringComparison.Ordinal)).Select(word => word.Text));
        Assert.Equal(45D, Assert.Single(words, word => word.Text == "ListLine00").Letters[0].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DecoratedNotesStillRejectEffectiveAnchoredContent(bool inherited) {
        PdfCreateOptions options = Options();
        var unsupported = new PdfPanelStyle {
            AnchoredCanvas = new PdfCanvasBlock(new PdfPageCanvas().ForegroundImage(PdfPngTestImages.Create16BitRgbPng(), 0, 0, 4, 4).Items)
        };
        if (inherited) options.DefaultPanelStyle = unsupported;
        var document = PdfDocument.Create(options);
        Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(body => body.Panel(
            panel => panel.Text("AnchoredNote", style: Style()), inherited ? null : unsupported)));
    }

    [Theory]
    [InlineData("Image", false)]
    [InlineData("Image", true)]
    [InlineData("Rule", false)]
    [InlineData("Rule", true)]
    public void NonTextNoteKeepsRetainTheirFollowingParagraph(string kind, bool inherited) {
        PdfCreateOptions options = Options();
        var imageStyle = new PdfImageStyle { KeepWithNext = true, SpacingBefore = 0, SpacingAfter = 0 };
        var ruleStyle = new PdfHorizontalRuleStyle { KeepWithNext = true, Thickness = 20, SpacingBefore = 0, SpacingAfter = 0 };
        if (inherited) {
            if (kind == "Image") options.DefaultImageStyle = imageStyle;
            else options.DefaultHorizontalRuleStyle = ruleStyle;
        }
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => {
            body.Text("NoteLead", style: NoteLineStyle(20));
            if (kind == "Image") body.Image(PdfPngTestImages.Create16BitRgbPng(), 4, 20, style: inherited ? null : imageStyle);
            else body.HR(style: inherited ? null : ruleStyle);
            body.Text("KeptCaption", style: NoteLineStyle(20));
        });
        document.Content.Spacer(50).Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("NoteLead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("KeptCaption", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("KeptCaption", pdf.GetPage(2).Text, StringComparison.Ordinal);
        if (kind == "Image") {
            Assert.Empty(pdf.GetPage(1).GetImages());
            Assert.Single(pdf.GetPage(2).GetImages());
        } else {
            Assert.Empty(pdf.GetPage(1).Paths);
            Assert.NotEmpty(pdf.GetPage(2).Paths);
        }
    }

    [Fact]
    public void NumberedNoteItemsStillRejectRecursiveReferencesBeforeTheyCanBeCaptured() {
        var document = PdfDocument.Create(Options());
        var original = document.Content.DefineFootnote(body => body.Text("OriginalNote", style: Style()));
        Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(body => body.RichNumbered(new[] {
            PdfListItem.Rich(new[] { PdfTextRun.Normal("Nested").WithFootnoteReference(original) })
        })));
    }

    [Theory]
    [InlineData(false, 2)]
    [InlineData(true, 1)]
    public void ParagraphPanelAfterSpacingUsesItsExistingFrameEndClippingPolicy(bool clip, int pages) {
        var document = PdfDocument.Create(Options());
        var note = document.Content.DefineFootnote(body => body.Panel(panel => panel.Text(
            "SpacingLine00\nSpacingLine01\nSpacingLine02\nSpacingLine03", style: NoteLineStyle(20)),
            new PdfPanelStyle {
                PaddingX = 0, PaddingY = 0, BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 30,
                ClipSpacingAfterAtFrameEnd = clip
            }));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(pages, pdf.NumberOfPages);
        Assert.Contains("SpacingLine03", pdf.GetPage(pages).Text, StringComparison.Ordinal);
        Assert.Equal(Enumerable.Range(0, 4).Select(index => $"SpacingLine{index:D2}"), pdf.GetPages()
            .SelectMany(page => page.GetWords()).Where(word => word.Text.StartsWith("SpacingLine", StringComparison.Ordinal))
            .Select(word => word.Text));
    }

    private static PdfParagraphStyle NoteLineStyle(double leading) => new() {
        FontSize = 8, LineSpacing = PdfLineSpacing.Exactly(leading), WidowControl = false,
        SpacingBefore = 0, SpacingAfter = 0
    };
}
