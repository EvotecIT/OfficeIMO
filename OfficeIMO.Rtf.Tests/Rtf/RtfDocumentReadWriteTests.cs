using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentReadWriteTests {
    [Fact]
    public void Read_Binds_Paragraphs_Runs_Formatting_And_Unicode() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}{\colortbl;\red255\green0\blue0;}\pard Hello {\b bold} {\i italic} \u380? \cf1 red\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);
        Assert.Equal(2, result.Document.Fonts.Count);
        Assert.Single(result.Document.Colors);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Contains("Hello bold italic ż red", paragraph.ToPlainText());
        Assert.Contains(paragraph.Runs, run => run.Text == "bold" && run.Bold);
        Assert.Contains(paragraph.Runs, run => run.Text == "italic" && run.Italic);
        Assert.Contains(paragraph.Runs, run => run.Text.Contains("red", StringComparison.Ordinal) && run.ForegroundColorIndex == 1);
    }

    [Fact]
    public void Read_Uses_Font_Charset_CodePage_For_Ansi_Bytes() {
        const string rtf = @"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fcharset0 Arial;}{\f1\fcharset238 Arial CE;}}\f1 Za\'bf\'f3\'b3\'e6 g\'ea\'9cl\'b9 ja\'9f\'f1\f0  \'bf\par}";

        RtfDocument document = RtfDocument.Read(rtf).Document;
        string text = Assert.Single(document.Paragraphs).ToPlainText();

        Assert.StartsWith("Zażółć gęślą jaźń", text, StringComparison.Ordinal);
        Assert.EndsWith("¿", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Read_Resolves_Large_Font_Tables_For_Repeated_Switches() {
        var rtf = new StringBuilder(@"{\rtf1\ansi\ansicpg1252{\fonttbl");
        for (int index = 0; index < 1_000; index++) {
            rtf.Append(@"{\f").Append(index).Append(index == 999 ? @"\fcharset238 " : @"\fcharset0 ")
                .Append("Font ").Append(index).Append(";}");
        }
        rtf.Append('}');
        for (int index = 0; index < 1_000; index++) rtf.Append(@"\f999 \'bf");
        rtf.Append('}');

        RtfDocument document = RtfDocument.Read(rtf.ToString()).Document;

        Assert.Equal(1_000, document.Fonts.Count);
        Assert.Equal(1_000, Assert.Single(document.Paragraphs).ToPlainText().Count(character => character == 'ż'));
    }

    [Fact]
    public void Write_Emits_Deterministic_Rtf_And_Reads_Back_Formatting() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Hello ");
        paragraph.AddText("RTF").SetBold().SetUnderline().SetFontSize(14);
        paragraph.AddText(" ż").ForegroundColorIndex = red;

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.StartsWith(@"{\rtf1\ansi\deff0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\fonttbl{\f0 Calibri;}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\colortbl;\red255\green0\blue0;}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\b \ul \fs28 RTF", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\u380?", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Hello RTF ż", readParagraph.ToPlainText());
        Assert.Contains(readParagraph.Runs, run => run.Text == "RTF" && run.Bold && run.Underline && run.FontSize == 14);
    }

    [Fact]
    public void Write_Provides_Encoded_Byte_And_Stream_Output() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Generated ż");
        var options = new RtfWriteOptions { IncludeGenerator = false };

        string expected = document.ToRtf(options);
        byte[] bytes = document.ToBytes(options);

        Assert.Equal(expected, Encoding.UTF8.GetString(bytes));

        using var stream = new MemoryStream();
        document.Save(stream, options);
        Assert.Equal(bytes, stream.ToArray());

        using MemoryStream memoryStream = document.ToStream(options);
        Assert.Equal(bytes, memoryStream.ToArray());

        RtfReadResult read = RtfDocument.Load(bytes);
        Assert.Equal("Generated ż", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public async Task Save_File_Defaults_To_Word_Compatible_Bomless_Utf8() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Word-compatible ż");
        string syncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        string asyncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");

        try {
            document.Save(syncPath);
            await document.SaveAsync(asyncPath);

            byte[] syncBytes = File.ReadAllBytes(syncPath);
            byte[] asyncBytes = File.ReadAllBytes(asyncPath);
            Assert.Equal((byte)'{', syncBytes[0]);
            Assert.Equal((byte)'{', asyncBytes[0]);
            Assert.False(syncBytes.Take(3).SequenceEqual(Encoding.UTF8.GetPreamble()));
            Assert.False(asyncBytes.Take(3).SequenceEqual(Encoding.UTF8.GetPreamble()));
            Assert.Equal("Word-compatible ż", Assert.Single(RtfDocument.Load(syncBytes).Document.Paragraphs).ToPlainText());
            Assert.Equal("Word-compatible ż", Assert.Single(RtfDocument.Load(asyncBytes).Document.Paragraphs).ToPlainText());
        } finally {
            File.Delete(syncPath);
            File.Delete(asyncPath);
        }
    }

    [Fact]
    public async Task Save_File_Preserves_Explicit_Encoding_Preambles() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Encoded ż");
        string syncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        string asyncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        Encoding syncEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        Encoding asyncEncoding = Encoding.Unicode;

        try {
            document.Save(syncPath, encoding: syncEncoding);
            await document.SaveAsync(asyncPath, encoding: asyncEncoding);

            AssertFileStartsWithPreamble(syncPath, syncEncoding);
            AssertFileStartsWithPreamble(asyncPath, asyncEncoding);
        } finally {
            File.Delete(syncPath);
            File.Delete(asyncPath);
        }
    }

    private static void AssertFileStartsWithPreamble(string path, Encoding encoding) {
        byte[] bytes = File.ReadAllBytes(path);
        byte[] preamble = encoding.GetPreamble();

        Assert.NotEmpty(preamble);
        Assert.True(bytes.Take(preamble.Length).SequenceEqual(preamble));
        Assert.StartsWith(@"{\rtf1", encoding.GetString(bytes, preamble.Length, bytes.Length - preamble.Length), StringComparison.Ordinal);
    }

    [Fact]
    public void Write_Clears_Sticky_Character_Table_Formatting_Between_Runs_And_Paragraphs() {
        RtfDocument document = RtfDocument.Create();
        int fontId = document.AddFont("Consolas");
        int red = document.AddColor(255, 0, 0);
        document.AddStyle(2, "Code", RtfStyleKind.Character).Additive = true;

        RtfRun paragraphRun = document.AddParagraph().AddText("Styled");
        paragraphRun.FontId = fontId;
        paragraphRun.FontSize = 18;
        paragraphRun.ForegroundColorIndex = red;
        paragraphRun.StyleId = 2;

        document.AddParagraph("Next");

        RtfParagraph mixed = document.AddParagraph();
        RtfRun mixedStyled = mixed.AddText("Again");
        mixedStyled.FontId = fontId;
        mixedStyled.FontSize = 18;
        mixedStyled.ForegroundColorIndex = red;
        mixedStyled.StyleId = 2;
        mixed.AddText(" plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"Styled\plain \par", rtf, StringComparison.Ordinal);
        Assert.Contains(@"Again\plain  plain", rtf, StringComparison.Ordinal);

        RtfRun nextRun = Assert.Single(read.Document.Paragraphs[1].Runs);
        Assert.Equal("Next", nextRun.Text);
        Assert.Null(nextRun.FontId);
        Assert.Null(nextRun.FontSize);
        Assert.Null(nextRun.ForegroundColorIndex);
        Assert.Null(nextRun.StyleId);

        RtfRun plainRun = read.Document.Paragraphs[2].Runs.Single(run => run.Text == " plain");
        Assert.Null(plainRun.FontId);
        Assert.Null(plainRun.FontSize);
        Assert.Null(plainRun.ForegroundColorIndex);
        Assert.Null(plainRun.StyleId);
    }

    [Fact]
    public void Read_Preserves_Syntax_And_Binds_Picture_Destinations() {
        const string rtf = @"{\rtf1\ansi{\pict\pngblip\bin3 abc}\pard Visible\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(result.SyntaxTree.Root.Children, node => node is OfficeIMO.Rtf.Syntax.RtfGroup group && group.Destination == "pict");
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.Severity == RtfDiagnosticSeverity.Warning);
        RtfImage image = Assert.IsType<RtfImage>(result.Document.Blocks[0]);
        Assert.Equal(RtfImageFormat.Png, image.Format);
        Assert.Equal(new byte[] { 97, 98, 99 }, image.Data);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Visible", paragraph.ToPlainText());
    }

    [Fact]
    public void Read_Uses_Ansi_CodePage_For_Hex_Escaped_Body_Text_And_Metadata() {
        const string rtf = @"{\rtf1\ansi\ansicpg1250{\info{\title \'a3\'b9\'8c}}\pard \'a3\'b9\'8c\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(RtfDocumentCharacterSet.Ansi, result.Document.Settings.CharacterSet);
        Assert.Equal(1250, result.Document.Settings.AnsiCodePage);
        Assert.Equal("ŁąŚ", result.Document.Info.Title);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("ŁąŚ", paragraph.ToPlainText());
    }

    [Theory]
    [InlineData(@"\mac", RtfDocumentCharacterSet.Mac, @"\'8e\'8f\'d0", "éè–")]
    [InlineData(@"\pc", RtfDocumentCharacterSet.Pc, @"\'9b\'9c\'9d", "¢£¥")]
    [InlineData(@"\pca", RtfDocumentCharacterSet.Pca, @"\'9b\'9c\'9d", "ø£Ø")]
    public void Read_Uses_Document_Character_Set_Default_For_Hex_Escaped_Text_And_Metadata(string characterSetControl, RtfDocumentCharacterSet expectedCharacterSet, string encodedText, string expectedText) {
        string rtf = $@"{{\rtf1{characterSetControl}{{\info{{\title {encodedText}}}}}\pard {encodedText}\par}}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(expectedCharacterSet, result.Document.Settings.CharacterSet);
        Assert.Null(result.Document.Settings.AnsiCodePage);
        Assert.Equal(expectedText, result.Document.Info.Title);
        Assert.Equal(expectedText, Assert.Single(result.Document.Paragraphs).ToPlainText());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF103");
    }

    [Fact]
    public void Read_Ansi_CodePage_Overrides_Document_Character_Set_Default() {
        const string rtf = @"{\rtf1\mac\ansicpg1252{\info{\title \'80}}\pard \'80\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(RtfDocumentCharacterSet.Mac, result.Document.Settings.CharacterSet);
        Assert.Equal(1252, result.Document.Settings.AnsiCodePage);
        Assert.Equal("€", result.Document.Info.Title);
        Assert.Equal("€", Assert.Single(result.Document.Paragraphs).ToPlainText());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF103");
    }


}
