using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfParagraphFormattingTests {
    [Fact]
    public void Read_Binds_Paragraph_Shading_And_Borders_Without_Leaking() {
        const string rtf = @"{\rtf1\ansi{\colortbl;\red230\green242\blue255;\red68\green114\blue196;\red0\green170\blue85;}\pard\cbpat1\brdrt\brdrs\brdrw12\brdrcf2\brdrl\brdrdb\brdrw8\brdrcf3 Boxed\par\pard Plain\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(2, result.Document.Paragraphs.Count);
        RtfParagraph boxed = result.Document.Paragraphs[0];
        Assert.Equal("Boxed", boxed.ToPlainText());
        Assert.Equal(1, boxed.BackgroundColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Single, boxed.TopBorder.Style);
        Assert.Equal(12, boxed.TopBorder.Width);
        Assert.Equal(2, boxed.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, boxed.LeftBorder.Style);
        Assert.Equal(8, boxed.LeftBorder.Width);
        Assert.Equal(3, boxed.LeftBorder.ColorIndex);

        RtfParagraph plain = result.Document.Paragraphs[1];
        Assert.Equal("Plain", plain.ToPlainText());
        Assert.Null(plain.BackgroundColorIndex);
        Assert.False(plain.TopBorder.HasAnyValue);
        Assert.False(plain.LeftBorder.HasAnyValue);
    }

    [Fact]
    public void Write_And_Read_Paragraph_Shading_And_Borders_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        int fill = document.AddColor(0xE6, 0xF2, 0xFF);
        int topColor = document.AddColor(0x44, 0x72, 0xC4);
        int leftColor = document.AddColor(0x00, 0xAA, 0x55);
        RtfParagraph boxed = document.AddParagraph("Boxed");
        boxed.SetBackgroundColor(fill)
            .SetBorder(RtfParagraphBorderSide.Top, RtfParagraphBorderStyle.Single, width: 12, colorIndex: topColor)
            .SetBorder(RtfParagraphBorderSide.Left, RtfParagraphBorderStyle.Double, width: 8, colorIndex: leftColor)
            .SetBorder(RtfParagraphBorderSide.Bottom, RtfParagraphBorderStyle.Dotted)
            .SetBorder(RtfParagraphBorderSide.Right, RtfParagraphBorderStyle.Dashed);
        document.AddParagraph("Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\cbpat1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrt\brdrs\brdrw12\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrl\brdrdb\brdrw8\brdrcf3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrb\brdrdot", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrr\brdrdash", rtf, StringComparison.Ordinal);

        RtfParagraph readBoxed = read.Document.Paragraphs[0];
        Assert.Equal(fill, readBoxed.BackgroundColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Single, readBoxed.TopBorder.Style);
        Assert.Equal(12, readBoxed.TopBorder.Width);
        Assert.Equal(topColor, readBoxed.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, readBoxed.LeftBorder.Style);
        Assert.Equal(8, readBoxed.LeftBorder.Width);
        Assert.Equal(leftColor, readBoxed.LeftBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Dotted, readBoxed.BottomBorder.Style);
        Assert.Equal(RtfParagraphBorderStyle.Dashed, readBoxed.RightBorder.Style);

        RtfParagraph readPlain = read.Document.Paragraphs[1];
        Assert.Null(readPlain.BackgroundColorIndex);
        Assert.False(readPlain.TopBorder.HasAnyValue);
        Assert.False(readPlain.LeftBorder.HasAnyValue);
        Assert.False(readPlain.BottomBorder.HasAnyValue);
        Assert.False(readPlain.RightBorder.HasAnyValue);
    }

    [Fact]
    public void Write_And_Read_Paragraph_Shading_Patterns_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        int fill = document.AddColor(0xE6, 0xF2, 0xFF);
        int foreground = document.AddColor(0x44, 0x72, 0xC4);
        RtfParagraph shaded = document.AddParagraph("Pattern");
        shaded.SetShading(fill, foregroundColorIndex: foreground, patternPercent: 3750, pattern: RtfShadingPattern.DarkForwardDiagonal);
        document.AddParagraph("Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\cbpat1\cfpat2\shading3750\bgdkfdiag", rtf, StringComparison.Ordinal);
        RtfParagraph readShaded = read.Document.Paragraphs[0];
        Assert.Equal(fill, readShaded.BackgroundColorIndex);
        Assert.Equal(foreground, readShaded.ShadingForegroundColorIndex);
        Assert.Equal(3750, readShaded.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, readShaded.ShadingPattern);

        RtfParagraph readPlain = read.Document.Paragraphs[1];
        Assert.Null(readPlain.BackgroundColorIndex);
        Assert.Null(readPlain.ShadingForegroundColorIndex);
        Assert.Null(readPlain.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.None, readPlain.ShadingPattern);
    }

    [Fact]
    public void Read_Binds_Paragraph_Shading_Patterns_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\colortbl;\red230\green242\blue255;\red68\green114\blue196;}\pard\cbpat1\cfpat2\shading6250\bgdcross Pattern\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Pattern", paragraph.ToPlainText());
        Assert.Equal(1, paragraph.BackgroundColorIndex);
        Assert.Equal(2, paragraph.ShadingForegroundColorIndex);
        Assert.Equal(6250, paragraph.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DiagonalCross, paragraph.ShadingPattern);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Paragraph_Shading_And_Borders() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Boxed");
        paragraph.ShadingFillColorHex = "e6f2ff";
        paragraph.Borders.TopStyle = BorderValues.Single;
        paragraph.Borders.TopSize = 12U;
        paragraph.Borders.TopColorHex = "4472c4";
        paragraph.Borders.LeftStyle = BorderValues.Double;
        paragraph.Borders.LeftSize = 8U;
        paragraph.Borders.LeftColorHex = "00aa55";

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal(1, rtfParagraph.BackgroundColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Single, rtfParagraph.TopBorder.Style);
        Assert.Equal(12, rtfParagraph.TopBorder.Width);
        Assert.Equal(2, rtfParagraph.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, rtfParagraph.LeftBorder.Style);
        Assert.Equal(8, rtfParagraph.LeftBorder.Width);
        Assert.Equal(3, rtfParagraph.LeftBorder.ColorIndex);
        Assert.Contains(@"\cbpat1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrt\brdrs\brdrw12\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrl\brdrdb\brdrw8\brdrcf3", rtf, StringComparison.Ordinal);

        WordParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs);
        Assert.Equal("e6f2ff", roundTripParagraph.ShadingFillColorHex);
        Assert.Equal(BorderValues.Single, roundTripParagraph.Borders.TopStyle);
        Assert.Equal(12U, roundTripParagraph.Borders.TopSize?.Value);
        Assert.Equal("4472c4", roundTripParagraph.Borders.TopColorHex);
        Assert.Equal(BorderValues.Double, roundTripParagraph.Borders.LeftStyle);
        Assert.Equal(8U, roundTripParagraph.Borders.LeftSize?.Value);
        Assert.Equal("00aa55", roundTripParagraph.Borders.LeftColorHex);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Applies_Paragraph_Shading_And_Borders() {
        RtfDocument document = RtfDocument.Create();
        int fill = document.AddColor(0xE6, 0xF2, 0xFF);
        int topColor = document.AddColor(0x44, 0x72, 0xC4);
        int leftColor = document.AddColor(0x00, 0xAA, 0x55);
        RtfParagraph paragraph = document.AddParagraph("Boxed");
        paragraph.SetBackgroundColor(fill)
            .SetBorder(RtfParagraphBorderSide.Top, RtfParagraphBorderStyle.Single, width: 12, colorIndex: topColor)
            .SetBorder(RtfParagraphBorderSide.Left, RtfParagraphBorderStyle.Double, width: 8, colorIndex: leftColor);

        using WordDocument word = document.ToWordDocument();

        WordParagraph wordParagraph = Assert.Single(word.Paragraphs);
        Assert.Equal("e6f2ff", wordParagraph.ShadingFillColorHex);
        Assert.Equal(BorderValues.Single, wordParagraph.Borders.TopStyle);
        Assert.Equal(12U, wordParagraph.Borders.TopSize?.Value);
        Assert.Equal("4472c4", wordParagraph.Borders.TopColorHex);
        Assert.Equal(BorderValues.Double, wordParagraph.Borders.LeftStyle);
        Assert.Equal(8U, wordParagraph.Borders.LeftSize?.Value);
        Assert.Equal("00aa55", wordParagraph.Borders.LeftColorHex);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Paragraph_Shading_Pattern_And_Foreground() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Pattern");
        paragraph.ShadingFillColorHex = "e6f2ff";
        paragraph.ShadingPattern = ShadingPatternValues.Percent37;
        paragraph._paragraph.ParagraphProperties!.Shading!.Color = "4472c4";

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal(1, rtfParagraph.BackgroundColorIndex);
        Assert.Equal(2, rtfParagraph.ShadingForegroundColorIndex);
        Assert.Equal(3750, rtfParagraph.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.None, rtfParagraph.ShadingPattern);
        Assert.Contains(@"\cbpat1\cfpat2\shading3750", rtf, StringComparison.Ordinal);

        WordParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs);
        Assert.Equal("e6f2ff", roundTripParagraph.ShadingFillColorHex);
        Assert.Equal("4472c4", roundTripParagraph._paragraphProperties?.Shading?.Color?.Value);
        Assert.Equal(ShadingPatternValues.Percent37, roundTripParagraph.ShadingPattern);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Applies_Paragraph_Shading_Pattern_And_Foreground() {
        RtfDocument document = RtfDocument.Create();
        int fill = document.AddColor(0xE6, 0xF2, 0xFF);
        int foreground = document.AddColor(0x44, 0x72, 0xC4);
        RtfParagraph paragraph = document.AddParagraph("Pattern");
        paragraph.SetShading(fill, foregroundColorIndex: foreground, pattern: RtfShadingPattern.DarkForwardDiagonal);

        using WordDocument word = document.ToWordDocument();

        WordParagraph wordParagraph = Assert.Single(word.Paragraphs);
        Assert.Equal("e6f2ff", wordParagraph.ShadingFillColorHex);
        Assert.Equal("4472c4", wordParagraph._paragraphProperties?.Shading?.Color?.Value);
        Assert.Equal(ShadingPatternValues.DiagonalStripe, wordParagraph.ShadingPattern);
    }
}
