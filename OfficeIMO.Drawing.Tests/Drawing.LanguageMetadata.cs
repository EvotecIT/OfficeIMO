using System;
using System.Globalization;
using System.Linq;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeLanguageMetadataTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Stacked_layout_keeps_language_on_every_measured_text_element(bool shrink) {
        var runs = new[] { new OfficeRichTextRun("ABC", 20, OfficeColor.Black).WithLanguageId(1033) };
        var layout = OfficeTextLayoutEngine.LayoutStackedRichTextBlock(runs, 30, shrink ? 30 : 100, 1.2,
            (value, size) => (value?.Length ?? 0) * size, shrinkToFit: shrink);
        var segments = layout.Lines.SelectMany(line => line.Segments).ToArray();
        Assert.Equal("ABC", string.Concat(segments.Select(segment => segment.Text)));
        Assert.All(segments, segment => Assert.Equal(1033U, segment.LanguageId));
        if (shrink) Assert.All(segments, segment => Assert.True(segment.FontSize < 20));
    }

    [Fact]
    public void Language_survives_fluent_drawing_tint_and_measured_segments_without_merging_languages() {
        var original = new OfficeRichTextRun("alpha", 12, OfficeColor.Black, fontFamily: "Arial");
        Assert.Null(original.LanguageId);
        Assert.Throws<ArgumentOutOfRangeException>(() => original.WithLanguageId(65536));
        OfficeRichTextRun english = original.WithLanguageId(1033).WithParagraphIndent(new OfficeTextParagraphIndent(2, 0))
            .WithTextCase(OfficeTextCase.Uppercase, CultureInfo.InvariantCulture);
        Assert.Equal("ALPHA", english.Text);
        Assert.Equal(1033U, english.LanguageId);
        OfficeRichTextRun french = new OfficeRichTextRun("beta", 12, OfficeColor.Black, fontFamily: "Arial").WithLanguageId(1036);
        var drawing = new OfficeDrawing(300, 100).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { english, french })
        }, 0, 0, 300, 100);
        OfficeDrawing copy = drawing.Clone();
        copy.ApplyColorTint(OfficeColor.Blue);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(copy.Elements));
        Assert.Equal(new uint?[] { 1033, 1036 }, text.Paragraphs.Single().Runs.Select(run => run.LanguageId));
        var layout = OfficeDrawingTextLayout.Create(text, 150, 50,
            (value, size, family, style) => (value?.Length ?? 0) * size / 2);
        var segments = layout.Lines.SelectMany(line => line.Segments).Where(segment => segment.Text.Length != 0).ToArray();
        Assert.Equal("ALPHAbeta", string.Concat(segments.Select(segment => segment.Text)));
        Assert.Equal(new uint?[] { 1033, 1036 }, segments.Select(segment => segment.LanguageId));
        Assert.Equal(1033U, english.LanguageId);
        Assert.Null(english.WithLanguageId(null).LanguageId);
    }
}
