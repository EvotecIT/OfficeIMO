using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCharacterAttributeReadingTests {
    [Theory]
    [InlineData("Regular", "Arial", false, false, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Normal)]
    [InlineData("Italic", "Arial", false, true, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Normal)]
    [InlineData("UnderlineSingle", "Arial", false, false, OfficeTextDecorationStyle.Single, false, OfficeTextBaseline.Normal)]
    [InlineData("UnderlineDouble", "Arial", false, false, OfficeTextDecorationStyle.Double, false, OfficeTextBaseline.Normal)]
    [InlineData("Strikethrough", "Arial", false, false, OfficeTextDecorationStyle.None, true, OfficeTextBaseline.Normal)]
    [InlineData("Superscript", "Arial", false, false, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Superscript)]
    [InlineData("Subscript", "Arial", false, false, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Subscript)]
    [InlineData("FontCambria", "Cambria", false, false, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Normal)]
    [InlineData("FontCourier", "Courier New", false, false, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Normal)]
    [InlineData("FontConsolas", "Consolas", false, false, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Normal)]
    [InlineData("ColorRed", "Arial", false, false, OfficeTextDecorationStyle.None, false, OfficeTextBaseline.Normal)]
    [InlineData("Combined", "Courier New", true, true, OfficeTextDecorationStyle.Double, true, OfficeTextBaseline.Normal)]
    public void Independent_native_attributes_reach_the_public_run(string name, string family, bool bold, bool italic,
        OfficeTextDecorationStyle underline, bool strike, OfficeTextBaseline baseline) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextCharacterAttributes/native-textcharacterattributes.pub"));
        Assert.Equal(2, document.Pages.Count); Assert.Equal(12, document.TextStories.Count);
        PublisherTextStory story = Assert.Single(document.TextStories, item => item.Text == "Native character control: " + name + "\n");
        OfficeRichTextRun run = Assert.Single(Assert.Single(story.Paragraphs).Runs);
        Assert.Equal(family, run.FontFamily); Assert.Equal(name == "Combined" ? 14 : 12, run.FontSize);
        Assert.Equal(bold, run.Bold); Assert.Equal(italic, run.Italic); Assert.Equal(underline, run.UnderlineStyle);
        Assert.Equal(strike ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None, run.StrikethroughStyle);
        Assert.Equal(baseline, run.Baseline);
        Assert.Equal(name == "ColorRed" ? OfficeColor.Red : name == "Combined" ? OfficeColor.Blue : OfficeColor.Black, run.Color);
    }

    [Fact]
    public void Native_underline_patterns_keep_representable_kinds_and_report_weight_length_and_compound_pattern_approximations() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextUnderlinePatterns/native-textunderlinepatterns.pub"));
        var expected = new[] { OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.Single, OfficeTextDecorationStyle.Words,
            OfficeTextDecorationStyle.Double, OfficeTextDecorationStyle.Dotted, OfficeTextDecorationStyle.Single,
            OfficeTextDecorationStyle.Dashed, OfficeTextDecorationStyle.Dashed, OfficeTextDecorationStyle.Dashed,
            OfficeTextDecorationStyle.Wavy, OfficeTextDecorationStyle.Wavy, OfficeTextDecorationStyle.Dotted,
            OfficeTextDecorationStyle.Dashed, OfficeTextDecorationStyle.Dashed, OfficeTextDecorationStyle.Dashed,
            OfficeTextDecorationStyle.Dashed, OfficeTextDecorationStyle.Dashed, OfficeTextDecorationStyle.Wavy };
        Assert.Equal(3, document.Pages.Count); Assert.Equal(18, document.TextStories.Count);
        for (int native = 0; native < expected.Length; native++) {
            PublisherTextStory story = Assert.Single(document.TextStories, item => item.Text == "Native character control: Underline" + native + "\n");
            Assert.Equal(expected[native], Assert.Single(Assert.Single(story.Paragraphs).Runs).UnderlineStyle);
        }
        Assert.Equal(new[] { 6, 9, 10, 16, 17, 18, 19, 20, 21, 22, 23 }, document.ReadReport.FidelityDiagnostics
            .Where(item => item.Code == "PUB_TEXT_UNDERLINE_APPROXIMATED" && item.LossKind == OfficeConversionLossKind.Approximation)
            .Select(item => int.Parse(item.Location!.Substring("Quill/underline/".Length))).Distinct().OrderBy(value => value));
    }
}
