using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfPigPage = UglyToad.PdfPig.Content.Page;

namespace OfficeIMO.Tests.Pdf;

public partial class RichParagraphWrappingTests {
    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_StyledSourceLeaderKeepsItsTypographyBeforeTextAndInlineContent(
        PdfTextWhitespaceMode whitespace, bool inline) {
        using var actual = PdfPigDocument.Open(RenderSourceTabTypography(whitespace, inline, literal: false));
        using var reference = PdfPigDocument.Open(RenderSourceTabTypography(whitespace, inline, literal: true));
        var page = Assert.Single(actual.GetPages());
        var referencePage = Assert.Single(reference.GetPages());
        var leaders = page.Letters.Where(letter => letter.Value == "_").ToArray();
        var referenceLeaders = referencePage.Letters.Where(letter => letter.Value == "_").ToArray();

        // The reference authors Courier24 glyphs directly. It has no tab-style
        // carrier and therefore does not share the source-tab merge path.
        Assert.True(referenceLeaders.Length >= 3);
        Assert.Equal(referenceLeaders.Length, leaders.Length);
        Assert.Equal(string.Concat(referencePage.Letters.Select(letter => letter.Value)),
            string.Concat(page.Letters.Select(letter => letter.Value)));
        for (int index = 0; index < leaders.Length; index++) {
            Assert.Contains("Courier", leaders[index].FontName, StringComparison.Ordinal);
            Assert.Equal(24D, leaders[index].FontSize, 3);
            Assert.Equal(referenceLeaders[index].StartBaseLine.X, leaders[index].StartBaseLine.X, 3);
            Assert.Equal(referenceLeaders[index].StartBaseLine.Y, leaders[index].StartBaseLine.Y, 3);
        }
        Assert.Equal(36D, leaders[0].StartBaseLine.X, 3);
        if (inline) {
            var box = Assert.Single(page.Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), rectangle => rectangle.HasValue)!.Value;
            var referenceBox = Assert.Single(referencePage.Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), rectangle => rectangle.HasValue)!.Value;
            Assert.Equal(36D + 72D, box.Left, 3);
            Assert.Equal(10D, box.Width, 3);
            Assert.Equal(10D, box.Height, 3);
            Assert.Equal(referenceBox.Bottom, box.Bottom, 3);
        } else {
            var following = page.Letters.First(letter => letter.Value == "X");
            Assert.Contains("Helvetica", following.FontName, StringComparison.Ordinal);
            Assert.Equal(12D, following.FontSize, 3);
            Assert.Equal(36D + 72D, following.StartBaseLine.X, 3);
            Assert.Equal(leaders[0].StartBaseLine.Y, following.StartBaseLine.Y, 3);
        }
        var next = Assert.Single(page.GetWords(), word => word.Text == "NEXT");
        var referenceNext = Assert.Single(referencePage.GetWords(), word => word.Text == "NEXT");
        Assert.Equal(36D, next.Letters[0].StartBaseLine.X, 3);
        Assert.True(next.Letters[0].StartBaseLine.Y < leaders[0].StartBaseLine.Y);
        Assert.Equal(referenceLeaders[0].StartBaseLine.Y - referenceNext.Letters[0].StartBaseLine.Y,
            leaders[0].StartBaseLine.Y - next.Letters[0].StartBaseLine.Y, 3);
        Assert.Equal(referenceNext.Letters[0].StartBaseLine.Y, next.Letters[0].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_LeaderNoneKeepsStyledTabGeometryWithoutAddingSourceFontHeight(
        PdfTextWhitespaceMode whitespace, bool inline) {
        // An explicit paragraph None overrides the run's requested Underscores.
        // Changing an inkless tab from Helvetica12 to Courier24 must not invent
        // a painted font box or move the following content.
        using var actual = PdfPigDocument.Open(RenderSourceTabTypography(whitespace, inline,
            literal: false, paragraphLeader: PdfTabLeaderStyle.None));
        using var control = PdfPigDocument.Open(RenderSourceTabTypography(whitespace, inline,
            literal: false, paragraphLeader: PdfTabLeaderStyle.None,
            sourceFont: PdfStandardFont.Helvetica, sourceSize: 12D));
        var page = Assert.Single(actual.GetPages());
        var controlPage = Assert.Single(control.GetPages());
        Assert.Equal(inline ? "NEXT" : "XNEXT", string.Concat(page.Letters.Select(letter => letter.Value)));
        Assert.Equal(controlPage.Letters.Count, page.Letters.Count);
        for (int index = 0; index < page.Letters.Count; index++) {
            Assert.Equal(controlPage.Letters[index].Value, page.Letters[index].Value);
            Assert.Equal(controlPage.Letters[index].FontName, page.Letters[index].FontName);
            Assert.Equal(controlPage.Letters[index].FontSize, page.Letters[index].FontSize, 3);
            Assert.Equal(controlPage.Letters[index].StartBaseLine.X, page.Letters[index].StartBaseLine.X, 3);
            Assert.Equal(controlPage.Letters[index].StartBaseLine.Y, page.Letters[index].StartBaseLine.Y, 3);
        }
        if (inline) {
            var box = Assert.Single(page.Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), rectangle => rectangle.HasValue)!.Value;
            var controlBox = Assert.Single(controlPage.Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), rectangle => rectangle.HasValue)!.Value;
            Assert.Equal(36D + 72D, box.Left, 3);
            Assert.Equal(controlBox.Left, box.Left, 3);
            Assert.Equal(controlBox.Bottom, box.Bottom, 3);
            Assert.Equal(controlBox.Width, box.Width, 3);
            Assert.Equal(controlBox.Height, box.Height, 3);
        } else {
            var following = page.Letters.First(letter => letter.Value == "X");
            Assert.Equal(36D + 72D, following.StartBaseLine.X, 3);
            Assert.Equal(12D, following.FontSize, 3);
        }
    }

    [Fact]
    public void ParagraphTabs_SubThresholdSourceLeaderDoesNotBorrowTheFollowingFont() {
        // The 30pt gap is below Courier24's three-glyph leader threshold.
        // Falling back to the following Helvetica12 style would paint a leader;
        // an inkless selected source leader must retain the plain-tab geometry.
        using var actual = PdfPigDocument.Open(RenderSourceTabTypography(
            PdfTextWhitespaceMode.Collapse, inline: false, literal: false, tabStop: 30D));
        using var control = PdfPigDocument.Open(RenderSourceTabTypography(
            PdfTextWhitespaceMode.Collapse, inline: false, literal: false,
            paragraphLeader: PdfTabLeaderStyle.None, tabStop: 30D));
        var page = Assert.Single(actual.GetPages());
        var controlPage = Assert.Single(control.GetPages());
        Assert.Equal("XNEXT", string.Concat(page.Letters.Select(letter => letter.Value)));
        Assert.Equal(controlPage.Letters.Count, page.Letters.Count);
        for (int index = 0; index < page.Letters.Count; index++) {
            Assert.Equal(controlPage.Letters[index].Value, page.Letters[index].Value);
            Assert.Equal(controlPage.Letters[index].FontName, page.Letters[index].FontName);
            Assert.Equal(controlPage.Letters[index].FontSize, page.Letters[index].FontSize, 3);
            Assert.Equal(controlPage.Letters[index].StartBaseLine.X, page.Letters[index].StartBaseLine.X, 3);
            Assert.Equal(controlPage.Letters[index].StartBaseLine.Y, page.Letters[index].StartBaseLine.Y, 3);
        }
        var following = page.Letters[0];
        Assert.Equal("X", following.Value);
        Assert.Contains("Helvetica", following.FontName, StringComparison.Ordinal);
        Assert.Equal(12D, following.FontSize, 3);
        Assert.Equal(36D + 30D, following.StartBaseLine.X, 3);
        var next = Assert.Single(page.GetWords(), word => word.Text == "NEXT");
        var controlNext = Assert.Single(controlPage.GetWords(), word => word.Text == "NEXT");
        Assert.Equal(controlNext.Letters[0].StartBaseLine.Y, next.Letters[0].StartBaseLine.Y, 3);
    }

    [Fact]
    public void ParagraphTabs_DirectContinuationMatchesBufferedSourceStyleAndTransfersOwnershipOnce() {
        var owner = PdfDocument.Create(TrailingTabOptions());
        PdfFootnoteDefinition note = owner.Content.DefineFootnote(body => body.Text("NOTE"));
        var fieldStyle = new PdfFormFieldStyle {
            BackgroundColor = null, BorderColor = null, BorderWidth = global::OfficeIMO.OfficeLength.Points(0D)
        };
        var field = new PdfInlineFormField(PdfCanvasFormFieldItem.Text(
            "AfterTab", "Y", "Y", 0, 0, 1, 12, 12, fieldStyle, fieldStyle));
        PdfTextRun source = new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D), font: PdfStandardFont.Courier,
            tabLeader: PdfTabLeaderStyle.Underscores)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 0))
            .WithHorizontalTextScaling(90D).WithCharacterSpacing(global::OfficeIMO.OfficeLength.Points(.3D));
        PdfTextRun linked = PdfTextRun.Link("X", "https://example.test/after-source-tab", fontSize: global::OfficeIMO.OfficeLength.Points(12D))
            .WithFootnoteReference(note);
        PdfTextRun formText = PdfTextRun.Normal("Y", fontSize: global::OfficeIMO.OfficeLength.Points(12D), font: PdfStandardFont.Helvetica)
            .WithFormField(field);
        var direct = RestoreSourceTabTypography(new[] { source, linked, formText });
        var buffered = RestoreSourceTabTypography(new[] {
            source, PdfTextRun.Normal(" ", fontSize: global::OfficeIMO.OfficeLength.Points(12D), font: PdfStandardFont.Helvetica), linked, formText
        });

        // This contrast exercises existing private continuation machinery. The
        // buffered space selects the established retained-source-style path; its
        // authored separator must not receive the link/form-field ownership.
        Assert.Equal(buffered.Height, direct.Height, 3);
        foreach (List<PdfTextRun> restored in new[] { direct.Runs, buffered.Runs }) {
            PdfTextRun tab = Assert.Single(restored, run => run.Text == "\t");
            Assert.Equal(source.Font, tab.Font);
            Assert.Equal((source.FontSize?.ToPoints()), (tab.FontSize?.ToPoints()));
            Assert.Equal(source.FontFamily, tab.FontFamily);
            Assert.Equal(source.Baseline, tab.Baseline);
            Assert.Equal(source.TabLeader, tab.TabLeader);
            Assert.Equal(source.TabAlignment, tab.TabAlignment);
            Assert.Equal(source.FeatureSettings, tab.FeatureSettings);
            Assert.Equal(source.HorizontalTextScaling, tab.HorizontalTextScaling);
            Assert.Equal((source.CharacterSpacing.ToPoints()), (tab.CharacterSpacing.ToPoints()));
            Assert.Null(tab.LinkUri);
            Assert.Null(tab.FormField);
            Assert.Null(tab.FootnoteReference);
            Assert.Single(restored, run => ReferenceEquals(run.FootnoteReference, note));
            PdfTextRun restoredLink = Assert.Single(restored, run => run.LinkUri == linked.LinkUri);
            Assert.Equal("X", restoredLink.Text);
            Assert.Equal(12D, (restoredLink.FontSize?.ToPoints()));
            PdfTextRun restoredField = Assert.Single(restored, run => ReferenceEquals(run.FormField, field));
            Assert.Equal("Y", restoredField.Text);
            Assert.Equal(12D, (restoredField.FontSize?.ToPoints()));
        }
        Assert.Null(source.FootnoteReference);
        Assert.Same(note, linked.FootnoteReference);
        Assert.Same(field, formText.FormField);
        Assert.Equal("https://example.test/after-source-tab", linked.LinkUri);
    }

    private static byte[] RenderSourceTabTypography(PdfTextWhitespaceMode whitespace, bool inline,
        bool literal, PdfTabLeaderStyle paragraphLeader = PdfTabLeaderStyle.Underscores,
        PdfStandardFont sourceFont = PdfStandardFont.Courier, double sourceSize = 24D, double tabStop = 72D) {
        var style = new PdfParagraphStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(12D), LineHeight = 1.2D, SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D),
            TextWhitespaceMode = whitespace
        };
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(tabStop), leader: paragraphLeader);
        var runs = new List<PdfTextRun> {
            literal
                ? PdfTextRun.Normal("_____", fontSize: global::OfficeIMO.OfficeLength.Points(24D), font: PdfStandardFont.Courier)
                : new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(sourceSize), font: sourceFont,
                    tabLeader: PdfTabLeaderStyle.Underscores)
        };
        runs.Add(inline
            ? PdfTextRun.Inline(new PdfInlineBox(global::OfficeIMO.OfficeLength.Points(10D), global::OfficeIMO.OfficeLength.Points(10D), background: global::OfficeIMO.OfficeColor.Black))
            : PdfTextRun.Normal("X", fontSize: global::OfficeIMO.OfficeLength.Points(12D), font: PdfStandardFont.Helvetica));
        return PdfDocument.Create(TrailingTabOptions())
            .Paragraph(paragraph => paragraph.Runs(runs), style: style)
            .Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("NEXT", fontSize: global::OfficeIMO.OfficeLength.Points(12D), font: PdfStandardFont.Helvetica)
            }), style: style)
            .ToBytes();
    }

    private static (List<PdfTextRun> Runs, double Height) RestoreSourceTabTypography(PdfTextRun[] runs) {
        object wrapped = InvokePrivateFontMethod<object>("WrapRichRunsCoreWithFirstLineOrigin",
            runs, 288D, 12D, PdfStandardFont.Helvetica, 14.4D, null!, null!, 36D,
            TrailingTabOptions(), new[] { new PdfTabStop(global::OfficeIMO.OfficeLength.Points(72D), leader: PdfTabLeaderStyle.Underscores) },
            null!, null!, null!, PdfTextWhitespaceMode.Collapse);
        object lines = wrapped.GetType().GetField("Item1")!.GetValue(wrapped)!;
        var restored = InvokePrivateFontMethod<List<PdfTextRun>>("BuildTextRunsFromWrappedLines", lines, 0, 1);
        var heights = (List<double>)wrapped.GetType().GetField("Item2")!.GetValue(wrapped)!;
        return (restored, Assert.Single(heights));
    }
}
