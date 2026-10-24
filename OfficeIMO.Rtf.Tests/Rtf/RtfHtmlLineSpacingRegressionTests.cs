using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlLineSpacingRegressionTests {
    [Theory]
    [InlineData("Implicit")]
    [InlineData("MultipleOnly")]
    [InlineData("Stylesheet")]
    public void ParsedRtf_ToHtml_ImplicitLineSpacing_PreservesNaturalLineHeight(string scenario) {
        string source = scenario switch {
            "MultipleOnly" => @"{\rtf1\ansi\pard\slmult1 Natural\line Continuation\par}",
            "Stylesheet" => @"{\rtf1\ansi{\stylesheet{\s0\sl480\slmult1 Normal;}}\pard\s0 Natural\line Continuation\par}",
            _ => @"{\rtf1\ansi\pard Natural\line Continuation\par}"
        };
        RtfDocument document = RtfDocument.Read(source).Document;
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Null(paragraph.LineSpacingTwips);

        string html = document.ToHtml();

        Assert.Contains("Natural", html, StringComparison.Ordinal);
        Assert.Contains("<br>", html, StringComparison.Ordinal);
        Assert.Contains("Continuation", html, StringComparison.Ordinal);
        Assert.DoesNotContain("line-height:", html, StringComparison.Ordinal);
        RtfParagraph effective = document.ResolveParagraphFormatting(paragraph);
        Assert.Equal(240, effective.LineSpacingTwips);
        Assert.True(effective.LineSpacingMultiple);
        Assert.Null(paragraph.LineSpacingTwips);
    }

    [Fact]
    public void ParsedRtf_ToHtml_ImplicitLineSpacing_RemainsNaturalInListsAndTableCells() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("List item").SetList(listId: 1, level: 0, kind: RtfListKind.Bullet);
        source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Cell text");
        RtfDocument document = RtfDocument.Read(source.ToRtf()).Document;

        string html = document.ToHtml();

        Assert.Contains("<li", html, StringComparison.Ordinal);
        Assert.Contains("List item", html, StringComparison.Ordinal);
        Assert.Contains("<table", html, StringComparison.Ordinal);
        Assert.Contains("Cell text", html, StringComparison.Ordinal);
        Assert.DoesNotContain("line-height:", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(240, true, "1")]
    [InlineData(360, true, "1.5")]
    [InlineData(360, false, "18pt")]
    [InlineData(360, null, "18pt")]
    public void ParsedRtf_ToHtml_ExplicitLineSpacing_PreservesRequestedCss(int twips, bool? multiple, string expectedLineHeight) {
        string multipleControl = multiple.HasValue ? (multiple.Value ? @"\slmult1" : @"\slmult0") : string.Empty;
        string source = @"{\rtf1\ansi\pard\sl" + twips + multipleControl + @" Explicit\par}";
        RtfDocument document = RtfDocument.Read(source).Document;
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);

        string html = document.ToHtml();

        Assert.Contains("line-height:" + expectedLineHeight + ";", html, StringComparison.Ordinal);
        Assert.Contains("Explicit", html, StringComparison.Ordinal);
        Assert.Equal(twips, paragraph.LineSpacingTwips);
        Assert.Equal(multiple, paragraph.LineSpacingMultiple);
    }

    [Fact]
    public void AuthoredRtf_ToHtml_InheritsStyleLineSpacing_WithoutAddingDirectFormatting() {
        RtfDocument document = RtfDocument.Create();
        RtfStyle style = document.AddStyle(7, "Spaced");
        style.LineSpacingTwips = 360;
        style.LineSpacingMultiple = true;
        RtfParagraph paragraph = document.AddParagraph("Styled text").SetStyle(style.Id);

        string html = document.ToHtml();

        Assert.Contains("line-height:1.5;", html, StringComparison.Ordinal);
        Assert.Contains("Styled text", html, StringComparison.Ordinal);
        Assert.Null(paragraph.LineSpacingTwips);
        Assert.Null(paragraph.LineSpacingMultiple);
    }
}
