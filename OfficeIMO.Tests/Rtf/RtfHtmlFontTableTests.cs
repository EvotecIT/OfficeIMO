using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlFontTableTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Rich_Font_Table_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfFont defaultFont = document.Fonts.Single(font => font.Id == 0);
        defaultFont.Family = RtfFontFamily.Swiss;
        defaultFont.Charset = 0;
        defaultFont.Pitch = 2;
        defaultFont.CodePage = 1252;
        defaultFont.Bias = 0;
        defaultFont.Panose = "020F0502020204030204";
        defaultFont.AlternateName = "Arial";
        defaultFont.NonTaggedName = "Calibri";
        defaultFont.Embedding = new RtfFontEmbedding {
            Type = RtfEmbeddedFontType.TrueType,
            FileName = "Calibri.ttf",
            FileCodePage = 1252,
            Data = new byte[] { 1, 2, 3, 255 }
        };
        int monospaceFontId = document.AddFont("Consolas");
        document.Settings.SetDefaultFont(monospaceFontId);
        RtfFont monospace = document.Fonts.Single(font => font.Id == monospaceFontId);
        monospace.Family = RtfFontFamily.Modern;
        monospace.Charset = 238;
        monospace.Pitch = 1;
        monospace.CodePage = 1250;
        document.AddParagraph().AddText("Code").FontId = monospaceFontId;

        string html = document.ToHtml(new RtfToHtmlOptions { FragmentOnly = false, NewLine = "\n" });

        Assert.Contains("<meta name=\"officeimo-rtf-fonts\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocument();
        Assert.Equal(monospaceFontId, roundTrip.Settings.DefaultFontId);
        RtfFont roundTripDefault = roundTrip.Fonts.Single(font => font.Id == 0);
        Assert.Equal(RtfFontFamily.Swiss, roundTripDefault.Family);
        Assert.Equal(0, roundTripDefault.Charset);
        Assert.Equal(2, roundTripDefault.Pitch);
        Assert.Equal(1252, roundTripDefault.CodePage);
        Assert.Equal(0, roundTripDefault.Bias);
        Assert.Equal("020F0502020204030204", roundTripDefault.Panose);
        Assert.Equal("Arial", roundTripDefault.AlternateName);
        Assert.Equal("Calibri", roundTripDefault.NonTaggedName);
        Assert.NotNull(roundTripDefault.Embedding);
        Assert.Equal(RtfEmbeddedFontType.TrueType, roundTripDefault.Embedding.Type);
        Assert.Equal("Calibri.ttf", roundTripDefault.Embedding.FileName);
        Assert.Equal(1252, roundTripDefault.Embedding.FileCodePage);
        Assert.Equal(new byte[] { 1, 2, 3, 255 }, roundTripDefault.Embedding.Data);
        RtfFont roundTripMonospace = roundTrip.Fonts.Single(font => font.Id == monospaceFontId);
        Assert.Equal("Consolas", roundTripMonospace.Name);
        Assert.Equal(RtfFontFamily.Modern, roundTripMonospace.Family);
        Assert.Equal(238, roundTripMonospace.Charset);
        Assert.Equal(1, roundTripMonospace.Pitch);
        Assert.Equal(1250, roundTripMonospace.CodePage);
        Assert.Equal(monospaceFontId, Assert.Single(Assert.Single(roundTrip.Paragraphs).Runs).FontId);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\fonttbl{\f0\fswiss\fcharset0\fprq2\cpg1252\fbias0{\*\panose 020F0502020204030204}{\*\fname Calibri}{\*\fontemb\fttruetype{\*\fontfile\cpg1252 Calibri.ttf} 010203ff} Calibri{\*\falt Arial};}{\f1\fmodern\fcharset238\fprq1\cpg1250 Consolas;}}", rtf, StringComparison.Ordinal);
    }
}
