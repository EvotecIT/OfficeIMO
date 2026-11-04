using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextHyphenationTests {
    private static double Measure(string? text, double size, string? family, OfficeFontStyle style) => (text?.Length ?? 0) * size / 2;
    private static string[] Lines(OfficeRichTextBlockLayout layout) => layout.Lines.Select(line =>
        string.Concat(line.Segments.Select(segment => segment.Text))).ToArray();
    private static OfficeRichTextParagraph Paragraph(string text, uint language = 1033, double zone = 0) {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 10, OfficeColor.Black).WithLanguageId(language) });
        paragraph.Hyphenation = OfficeRichTextHyphenation.Create(paragraph.Runs, zone, default);
        return paragraph;
    }

    [Theory]
    [InlineData(1033U, "representation", 30D, "repre-|senta-|tion", new[] { 5, 10, 14 })]
    [InlineData(9U, "representation", 30D, "repre-|senta-|tion", new[] { 5, 10, 14 })]
    [InlineData(1031U, "Silbentrennung", 25D, "Sil-|ben-|tren-|nung", new[] { 3, 6, 10, 14 })]
    [InlineData(7U, "Silbentrennung", 25D, "Sil-|ben-|tren-|nung", new[] { 3, 6, 10, 14 })]
    [InlineData(1031U, "BA\u0308CKEREI", 20D, "BA\u0308-|CKE-|REI", new[] { 3, 6, 9 })]
    [InlineData(2057U, "representation", 30D, "repres|entati|on", new[] { 6, 12, 14 })]
    public void Shared_rich_text_uses_supported_original_word_points_and_keeps_source_ends(uint language, string text,
        double width, string expected, int[] ends) {
        OfficeRichTextParagraph paragraph = Paragraph(text, language);
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, width, 200, Measure);
        Assert.Equal(expected.Split('|'), Lines(layout));
        Assert.Equal(ends, layout.Lines.Select(line => line.SourceTextEnd));
        Assert.Equal(text, string.Concat(paragraph.Runs.Select(run => run.Text)));
        Assert.All(layout.Lines, line => Assert.True(line.Width <= width));
    }

    [Theory]
    [InlineData(0D, "A ex-|tra-|ordi-|nary")]
    [InlineData(18D, "A |ex-|tra-|ordi-|nary")]
    public void Zone_moves_a_word_before_hyphenating_when_the_remaining_gap_is_small(double zone, string expected) {
        var paragraph = Paragraph("A extraordinary", zone: zone);
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 28, 200, Measure);
        Assert.Equal(expected.Split('|'), Lines(layout));
        Assert.Equal(15, layout.Lines.Last().SourceTextEnd);
    }

    [Theory]
    [InlineData(true, true, "repres|entati|on")]
    [InlineData(false, false, "representation")]
    public void Suppression_and_no_wrap_keep_their_existing_source_and_display_contracts(bool suppressed, bool wrap, string expected) {
        var paragraph = Paragraph("representation").WithAutomaticHyphenationSuppressed(suppressed);
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 30, 200, Measure, wrap: wrap);
        Assert.Equal(expected.Split('|'), Lines(layout));
        Assert.DoesNotContain(layout.Lines, line => line.EndsWithGeneratedHyphen);
        Assert.Equal(14, layout.Lines.Last().SourceTextEnd);
    }

    [Theory]
    [InlineData(1033U, "repre-|senta-|tion", "pre-")]
    [InlineData(1031U, "re|presen|tation", "re")]
    public void Styled_words_keep_suffix_style_and_do_not_invent_a_dictionary_across_languages(uint secondLanguage, string expected, string lastPrefix) {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("re", 10, OfficeColor.Red).WithLanguageId(1033),
            new OfficeRichTextRun("presentation", 10, OfficeColor.Blue, bold: true).WithLanguageId(secondLanguage)
        });
        paragraph.Hyphenation = OfficeRichTextHyphenation.Create(paragraph.Runs, 0, default);
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 30, 200, Measure);
        Assert.Equal(expected.Split('|'), Lines(layout));
        Assert.Equal(secondLanguage == 1033 ? 5 : 2, layout.Lines[0].SourceTextEnd);
        Assert.Equal(secondLanguage == 1033, layout.Lines[0].EndsWithGeneratedHyphen);
        Assert.Equal(OfficeColor.Red, layout.Lines[0].Segments[0].Color);
        var suffix = layout.Lines[0].Segments.Last();
        Assert.Equal(lastPrefix, suffix.Text); Assert.Equal(secondLanguage == 1033, suffix.Bold);
        Assert.Equal(secondLanguage == 1033 ? OfficeColor.Blue : OfficeColor.Red, suffix.Color);
        Assert.Equal(new[] { "re", "presentation" }, paragraph.Runs.Select(run => run.Text));
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(2D)]
    public void Region_slices_repaint_suffixes_and_keep_original_dictionary_context_at_different_widths(double scale) {
        var paragraph = Paragraph("obligatory");
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
        var source = new List<string>();
        foreach (var region in new (double Width, string Painted, int End)[] { (30D, "oblig-", 5), (10D, "a-", 6), (40D, "tory", 10) }) {
            var selected = Assert.Single(flow.Take(region.Width, 12, Measure, default));
            source.Add(string.Concat(selected.Runs.Select(run => run.Text)));
            Assert.Equal(region.End, flow.CharacterPosition);
            var painted = OfficeDrawingTextLayout.CreateParagraphs(new[] { selected }, region.Width * scale, 12 * scale, Measure, scale: scale);
            Assert.Equal(new[] { region.Painted }, Lines(painted));
            Assert.Equal(region.Painted.EndsWith("-", StringComparison.Ordinal), Assert.Single(painted.Lines).EndsWithGeneratedHyphen);
        }
        Assert.False(flow.HasRemaining);
        Assert.Equal("obligatory", string.Concat(source));
        Assert.Equal("obligatory", Assert.Single(paragraph.Runs).Text);
    }
}
