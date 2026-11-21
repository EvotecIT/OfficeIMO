using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordFontPhysicalLengthsTests {
    [Fact]
    public void ParagraphFontRoundsToNativeHalfPointsAndNullRemovesTheOverride() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Native font size");
        paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(10.25);
        Assert.Equal("21", paragraph._run!.RunProperties!.FontSize!.Val!.Value);
        using var stream = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(stream);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(10.5), reopened.Paragraphs[0].FontSize);
        reopened.Paragraphs[0].FontSize = null;
        using var clearedStream = new MemoryStream(reopened.ToBytes());
        using WordDocument cleared = WordDocument.Load(clearedStream);
        Assert.Null(cleared.Paragraphs[0].FontSize);
        Assert.Null(cleared.Paragraphs[0]._run!.RunProperties?.FontSize);
    }

    [Fact]
    public void ParagraphStyleWritesNativeTwipsAndHalfPointsAndReadsThemAsLengths() {
        using WordDocument document = WordDocument.Create();
        WordStyle style = document.Styles.AddParagraphStyle("PhysicalStyle");
        style.FontSize = global::OfficeIMO.OfficeLength.Points(11.25);
        style.IndentationBefore = global::OfficeIMO.OfficeLength.Points(-1.03);
        style.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(2.03);
        style.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        var native = style.Element;
        Assert.Equal("23", native.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal("-21", native.StyleParagraphProperties!.Indentation!.Left!.Value);
        Assert.Equal("41", native.StyleParagraphProperties.SpacingBetweenLines!.Before!.Value);
        Assert.Equal("0", native.StyleParagraphProperties.SpacingBetweenLines.After!.Value);
        document.AddParagraph("Styled text").SetStyle(style);
        using var stream = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(stream);
        WordStyle readback = reopened.Styles["PhysicalStyle"];
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(11.5), readback.FontSize);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(-1.05), readback.IndentationBefore);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(2.05), readback.LineSpacingBefore);
        Assert.Equal(global::OfficeIMO.OfficeLength.Zero, readback.LineSpacingAfter);
    }
}
