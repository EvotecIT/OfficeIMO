using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentReadWriteTests {
    [Fact]
    public void Load_Decodes_Literal_Ansi_Text_With_CodePage_And_Preserves_Source_Bytes() {
        var bytes = new List<byte>();
        bytes.AddRange(Encoding.ASCII.GetBytes(@"{\rtf1\ansi\ansicpg1250{\fonttbl{\f0 "));
        bytes.AddRange(new byte[] { 0xA3, 0xB9, 0x8C });
        bytes.AddRange(Encoding.ASCII.GetBytes(@";}}{\stylesheet{\s1 "));
        bytes.AddRange(new byte[] { 0xA3, 0xB9, 0x8C });
        bytes.AddRange(Encoding.ASCII.GetBytes(@";}}{\info{\title "));
        bytes.AddRange(new byte[] { 0xA3, 0xB9, 0x8C });
        bytes.AddRange(Encoding.ASCII.GetBytes(@"}}\pard "));
        bytes.AddRange(new byte[] { 0xA3, 0xB9, 0x8C });
        bytes.AddRange(Encoding.ASCII.GetBytes(@"\par}"));
        byte[] source = bytes.ToArray();

        using var input = new MemoryStream(source);
        RtfReadResult result = RtfDocument.Load(input);
        using var output = new MemoryStream();
        result.SaveLossless(output);

        Assert.Equal(source, output.ToArray());
        Assert.Equal("ŁąŚ", result.Document.Fonts[0].Name);
        Assert.Equal("ŁąŚ", Assert.Single(result.Document.Styles).Name);
        Assert.Equal("ŁąŚ", result.Document.Info.Title);
        Assert.Equal("ŁąŚ", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    public static IEnumerable<object[]> CharacterSetLiteralData() {
        yield return new object[] { @"\mac", RtfDocumentCharacterSet.Mac, new byte[] { 0x8E, 0x8F, 0xD0 }, "éè–" };
        yield return new object[] { @"\pc", RtfDocumentCharacterSet.Pc, new byte[] { 0x9B, 0x9C, 0x9D }, "¢£¥" };
        yield return new object[] { @"\pca", RtfDocumentCharacterSet.Pca, new byte[] { 0x9B, 0x9C, 0x9D }, "ø£Ø" };
    }

    [Theory]
    [MemberData(nameof(CharacterSetLiteralData))]
    public void Load_Decodes_Literal_Text_With_Document_Character_Set_And_Preserves_Source_Bytes(string characterSetControl, RtfDocumentCharacterSet expectedCharacterSet, byte[] encodedBytes, string expectedText) {
        var bytes = new List<byte>();
        bytes.AddRange(Encoding.ASCII.GetBytes("{\\rtf1" + characterSetControl + "{\\info{\\title "));
        bytes.AddRange(encodedBytes);
        bytes.AddRange(Encoding.ASCII.GetBytes("}}\\pard "));
        bytes.AddRange(encodedBytes);
        bytes.AddRange(Encoding.ASCII.GetBytes("\\par}"));
        byte[] source = bytes.ToArray();

        using var input = new MemoryStream(source);
        RtfReadResult result = RtfDocument.Load(input);
        using var output = new MemoryStream();
        result.SaveLossless(output);

        Assert.Equal(source, output.ToArray());
        Assert.Equal(expectedCharacterSet, result.Document.Settings.CharacterSet);
        Assert.Null(result.Document.Settings.AnsiCodePage);
        Assert.Equal(expectedText, result.Document.Info.Title);
        Assert.Equal(expectedText, Assert.Single(result.Document.Paragraphs).ToPlainText());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF103");
    }

    [Fact]
    public void Read_Uses_Uc_Control_When_Collecting_Metadata_And_Field_Text() {
        const string rtf = @"{\rtf1\ansi{\info{\title \uc2\u380?? Title}}\pard {\field{\*\fldinst HYPERLINK ""https://example.test""}{\fldrslt \uc2\u380?? Link}}\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal("ż Title", result.Document.Info.Title);
        AssertSingleHyperlinkField(Assert.Single(result.Document.Paragraphs), "ż Link");
    }

    [Fact]
    public void Read_Scopes_Uc_Control_When_Collecting_Nested_Metadata_And_Field_Text() {
        const string rtf = @"{\rtf1\ansi{\info{\title {\uc2\u380?? Nested} \u380?! After}}\pard {\field{\*\fldinst HYPERLINK ""https://example.test""}{\fldrslt {\uc2\u380?? Nested} \u380?! Link}}\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal("ż Nested ż! After", result.Document.Info.Title);
        AssertSingleHyperlinkField(Assert.Single(result.Document.Paragraphs), "ż Nested ż! Link");
    }

    [Fact]
    public void Read_Uses_Unicode_Alternative_From_Upr_Groups_For_Body_Metadata_And_Fields() {
        const string rtf = @"{\rtf1\ansi{\info{\title {\upr{Fallback title}{\*\ud{\u380? title}}}}}\pard {\upr{Fallback body}{\*\ud{\u380? body}}}\par{\field{\*\fldinst HYPERLINK ""https://example.test""}{\fldrslt {\upr{Fallback link}{\*\ud{\u380? link}}}}}}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal("ż title", result.Document.Info.Title);
        Assert.Equal("ż body", result.Document.Paragraphs[0].ToPlainText());
        AssertSingleHyperlinkField(result.Document.Paragraphs[1], "ż link");
    }

    [Fact]
    public void Read_Combines_Unicode_Surrogate_Pairs_In_Body_Metadata_And_Fields() {
        const string rtf = @"{\rtf1\ansi{\info{\title Smile \u-10179?\u-8704?}}\pard Body \u-10179?\u-8704?\par{\field{\*\fldinst HYPERLINK ""https://example.test""}{\fldrslt Link \u-10179?\u-8704?}}}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal("Smile 😀", result.Document.Info.Title);
        Assert.Equal("Body 😀", result.Document.Paragraphs[0].ToPlainText());
        AssertSingleHyperlinkField(result.Document.Paragraphs[1], "Link 😀");
    }

    [Fact]
    public void Read_Binds_Named_Special_Character_Controls_In_Body_Metadata_And_Fields() {
        const string rtf = @"{\rtf1\ansi{\info{\title A\endash B\emdash C \lquote quote\rquote  \ldblquote double\rdblquote  \bullet}}\pard A\endash B\emdash C \lquote quote\rquote  \ldblquote double\rdblquote  \bullet\par{\field{\*\fldinst HYPERLINK ""https://example.test""}{\fldrslt Link\endash \bullet}}}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal("A–B—C ‘quote’ “double” •", result.Document.Info.Title);
        Assert.Equal("A–B—C ‘quote’ “double” •", result.Document.Paragraphs[0].ToPlainText());
        AssertSingleHyperlinkField(result.Document.Paragraphs[1], "Link–•");
    }

    [Fact]
    public void Write_Emits_Special_Character_Controls_And_Reads_Them_Back() {
        const string specialText = "A\u2013B\u2014C \u2018quote\u2019 \u201Cdouble\u201D \u2022 \u00A0\u2011\u00AD\u2003\u2002\u2005\u200E\u200F\u200D\u200C";
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = specialText;
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Before ");
        paragraph.AddText(specialText).SetBold();
        RtfField field = paragraph.AddField("PAGE");
        field.AddText(specialText).SetItalic();

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\endash", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\emdash", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\lquote", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\rquote", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ldblquote", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\rdblquote", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\bullet", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\~", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\_", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\-", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\emspace", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\enspace", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\qmspace", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ltrmark", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\rtlmark", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\zwj", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\zwnj", rtf, StringComparison.Ordinal);
        Assert.Equal(specialText, result.Document.Info.Title);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Before " + specialText + specialText, readParagraph.ToPlainText());
        Assert.Equal(specialText, string.Concat(readParagraph.Runs.Where(run => run.Bold).Select(run => run.Text)));
        RtfField readField = Assert.Single(readParagraph.Inlines.OfType<RtfField>());
        Assert.Equal(specialText, readField.ToPlainText());
        Assert.Equal(specialText, string.Concat(readField.Result.Runs.Where(run => run.Italic).Select(run => run.Text)));
    }

    private static RtfField AssertSingleHyperlinkField(RtfParagraph paragraph, string expectedText) {
        RtfField field = Assert.Single(paragraph.Inlines.OfType<RtfField>());
        Assert.Equal(expectedText, field.ToPlainText());
        Assert.Equal(new Uri("https://example.test"), field.Hyperlink);
        return field;
    }
}
