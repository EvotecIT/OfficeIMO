using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfTextDecodingRegressionTests {
    [Fact]
    public void Picture_Paragraph_Terminator_Does_Not_Consume_A_Subsequent_Empty_Paragraph() {
        RtfDocument document = RtfDocument.Read(@"{\rtf1\ansi{\pict\pngblip 89504e47}\par\par End\par}").Document;
        Assert.IsType<RtfImage>(document.Blocks[0]);
        Assert.Equal(new[] { "", "End" }, document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
    }

    public static IEnumerable<object[]> TextCases() {
        yield return new object[] { @"{\rtf1\ansi\uc1 \u945\{X\par}", "αX" };
        yield return new object[] { @"{\rtf1\ansi\uc1 \u945\emdash X\par}", "αX" };
        yield return new object[] { @"{\rtf1\ansi\uc1 \u945{\b X}Y\par}", "αXY" };
        yield return new object[] { "{\\rtf1\\ansi Hello\r\nWorld\\par}", "HelloWorld" };
        yield return new object[] {
            @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fcharset0 Arial;}{\f1\fcharset204 Arial;}}\f1\'c0\par\pard\'c1\par}",
            "А|Б"
        };
        yield return new object[] { @"{\rtf1\ansi\ansicpg932 " + "\u0082" + @"\'a0\par}", "あ" };
        yield return new object[] { @"{\rtf1\ansi\ansicpg932 \'82" + "\u00a0" + @"\par}", "あ" };
        yield return new object[] { @"{\rtf1\ansi\uc2 \u945\'61\'62X\par}", "αX" };
    }

    [Theory]
    [MemberData(nameof(TextCases))]
    public void Read_Decodes_Source_Controls_And_Byte_Forms_Consistently(string rtf, string expected) {
        RtfDocument document = RtfDocument.Read(rtf).Document;
        Assert.Equal(expected, string.Join("|", document.Paragraphs.Select(paragraph => paragraph.ToPlainText())));
        RtfDocument normalized = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal(expected, string.Join("|", normalized.Paragraphs.Select(paragraph => paragraph.ToPlainText())));
    }

    [Fact]
    public void Read_Preserves_Explicit_Empty_Paragraphs_Without_An_Eof_Paragraph() {
        RtfDocument document = RtfDocument.Read(@"{\rtf1\ansi A\par\par B\par}").Document;
        Assert.Equal(new[] { "A", "", "B" }, document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        RtfDocument normalized = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal(new[] { "A", "", "B" }, normalized.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
    }

    [Fact]
    public void Read_Finalizes_Last_Paragraph_With_Active_Formatting() {
        RtfParagraph paragraph = Assert.Single(RtfDocument.Read(@"{\rtf1\ansi\pard\qc\li720 Center}").Document.Paragraphs);
        Assert.Equal("Center", paragraph.ToPlainText());
        Assert.Equal(RtfTextAlignment.Center, paragraph.Alignment);
        Assert.Equal(720, paragraph.LeftIndentTwips);
    }

    [Theory]
    [InlineData(@"{\rtf1\ansi{\info{\title \uc1\u945\{X}}Body\par}", "αX")]
    [InlineData(@"{\rtf1\ansi{\info{\title \uc1\u945{\b X}Y}}Body\par}", "αXY")]
    [InlineData("{\\rtf1\\ansi{\\info{\\title Hello\r\nWorld}}Body\\par}", "HelloWorld")]
    public void Read_Metadata_Uses_The_Same_Fallback_And_Trivia_Rules(string rtf, string expected) {
        Assert.Equal(expected, RtfDocument.Read(rtf).Document.Info.Title);
    }
}
