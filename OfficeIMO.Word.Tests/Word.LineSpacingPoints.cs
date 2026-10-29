using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("auto", 240)]
    [InlineData("auto", 360)]
    [InlineData("exact", 361)]
    [InlineData("atLeast", 362)]
    [InlineData("auto", 0)]
    [InlineData("exact", 0)]
    [InlineData("atLeast", 0)]
    public void MixedSpacing_SavesTheNativePairAndReopens(string rule, int value) {
        var expected = rule == "auto" ? WordLineSpacing.Multiple(value / 240D)
            : rule == "exact" ? WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(value * 635L))
            : WordLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Emus(value * 635L));
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Mixed spacing");
        paragraph.LineSpacing = expected;
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(10);
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(12);
        byte[] bytes = document.ToBytes();
        using (var package = WordprocessingDocument.Open(new MemoryStream(bytes), false)) {
            var spacing = package.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First().ParagraphProperties!.SpacingBetweenLines!;
            Assert.Equal(value.ToString(System.Globalization.CultureInfo.InvariantCulture), spacing.Line!.Value);
            Assert.Equal(rule, spacing.LineRule!.InnerText);
            Assert.Equal("200", spacing.Before!.Value);
            Assert.Equal("240", spacing.After!.Value);
        }
        using var reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal(expected, reopened.Paragraphs[0].LineSpacing);
        Assert.Equal(expected, Assert.Single(Assert.Single(reopened.Inspect().Sections).Elements.OfType<WordParagraphInfo>()).LineSpacing);
    }

    [Fact]
    public void MixedSpacing_OmittedRuleIsAMultipleAndReadDoesNotAuthorARule() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Automatic");
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines { Line = "240" };
        Assert.Equal(WordLineSpacing.Multiple(1), paragraph.LineSpacing);
        Assert.Null(paragraph.LineSpacing!.Value.Height);
        Assert.Null(paragraph._paragraph.ParagraphProperties.SpacingBetweenLines.LineRule);
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(WordLineSpacing.Multiple(1), reopened.Paragraphs[0].LineSpacing);
        Assert.Null(reopened.Paragraphs[0]._paragraph.ParagraphProperties!.SpacingBetweenLines!.LineRule);
    }

    [Theory]
    [InlineData(null, "atLeast")]
    [InlineData("garbage", "exact")]
    [InlineData("-1", "auto")]
    [InlineData("2147483648", "auto")]
    [InlineData("240", "futureRule")]
    public void MixedSpacing_UnsupportedNativeDeclarationsRemainUntouched(string? value, string rule) {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Native preservation");
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        var spacing = new SpacingBetweenLines { Line = value, Before = "200", After = "240" };
        spacing.SetAttribute(new OpenXmlAttribute("w", "lineRule", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", rule));
        paragraph._paragraph.ParagraphProperties.SpacingBetweenLines = spacing;
        string original = spacing.OuterXml;
        Assert.Null(paragraph.LineSpacing);
        Assert.Equal(original, spacing.OuterXml);
        byte[] bytes = document.ToBytes();
        using var reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Null(reopened.Paragraphs[0].LineSpacing);
        Assert.Equal(original, reopened.Paragraphs[0]._paragraph.ParagraphProperties!.SpacingBetweenLines!.OuterXml);
        paragraph.LineSpacing = null;
        Assert.Null(spacing.Line);
        Assert.Null(spacing.LineRule);
        Assert.Equal("200", spacing.Before!.Value);
        Assert.Equal("240", spacing.After!.Value);
    }

    [Fact]
    public void MixedSpacing_RoundsAtTheNativeWriteBoundaryAndClearsAtomically() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Rounding");
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines {
            Before = "200", After = "240", BeforeLines = 100, AfterLines = 200,
            BeforeAutoSpacing = true, AfterAutoSpacing = false
        };
        paragraph.LineSpacing = WordLineSpacing.Multiple(1.0020833333333334);
        Assert.Equal(241, paragraph.NativeLineSpacingValue);
        paragraph.LineSpacing = WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(953));
        Assert.Equal(2, paragraph.NativeLineSpacingValue);
        paragraph.LineSpacing = WordLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Zero);
        Assert.Equal(WordLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Zero), paragraph.LineSpacing);
        paragraph.LineSpacing = null;
        Assert.Null(paragraph.NativeLineSpacingValue);
        Assert.Null(paragraph.NativeLineSpacingRule);
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        var spacing = reopened.Paragraphs[0]._paragraph.ParagraphProperties!.SpacingBetweenLines!;
        Assert.Null(spacing.Line);
        Assert.Null(spacing.LineRule);
        Assert.Equal("200", spacing.Before!.Value);
        Assert.Equal("240", spacing.After!.Value);
        Assert.Equal(100, spacing.BeforeLines!.Value);
        Assert.Equal(200, spacing.AfterLines!.Value);
        Assert.True(spacing.BeforeAutoSpacing!.Value);
        Assert.False(spacing.AfterAutoSpacing!.Value);
        Assert.Throws<System.ArgumentOutOfRangeException>(() => WordLineSpacing.Multiple(double.NaN));
        Assert.Throws<System.ArgumentOutOfRangeException>(() => WordLineSpacing.Multiple(-1));
        Assert.Throws<System.ArgumentOutOfRangeException>(() => WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(-1)));
    }

    [Fact]
    public void MixedSpacing_DefaultAuthorsZeroWhileNullLeavesSpacingAbsent() {
        using var document = WordDocument.Create();
        var absent = document.AddParagraph("Absent");
        absent.LineSpacing = null;
        Assert.Null(absent._paragraph.ParagraphProperties?.SpacingBetweenLines);
        var zero = document.AddParagraph("Explicit zero");
        zero.LineSpacing = default(WordLineSpacing);
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Null(reopened.Paragraphs[0].LineSpacing);
        Assert.Equal(WordLineSpacing.Multiple(0), reopened.Paragraphs[1].LineSpacing);
        var spacing = reopened.Paragraphs[1]._paragraph.ParagraphProperties!.SpacingBetweenLines!;
        Assert.Equal("0", spacing.Line!.Value);
        Assert.Equal("auto", spacing.LineRule!.InnerText);
    }
}
