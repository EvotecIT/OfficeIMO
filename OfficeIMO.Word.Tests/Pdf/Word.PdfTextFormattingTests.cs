using System.Collections.Generic;
using System.Reflection;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfTextFormattingTests {
    [Fact]
    public void NativeRunFormattingProjectsToTypedPdfRuns() {
        using WordDocument document = WordDocument.Create();
        WordParagraph authored = document.AddParagraph("Styled");
        authored.Bold = true;
        authored.Italic = true;
        authored.Underline = WordUnderlineStyle.Dotted;
        authored.DoubleStrike = true;
        authored.VerticalTextAlignment = WordVerticalTextPosition.Superscript;
        authored.CapsStyle = WordCapsStyle.Caps;
        authored.ColorHex = "336699";
        authored.FontFamily = "Aptos";
        authored.FontSizePoints = 14D;

        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod(
            "CreateNativeCellParagraphRuns",
            BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            new[] { typeof(WordParagraph), typeof(Dictionary<long, int>) },
            modifiers: null)!;
        var runs = Assert.IsAssignableFrom<IReadOnlyList<PdfTextRun>>(
            method.Invoke(null, new object?[] { authored, null }));
        PdfTextRun run = Assert.Single(runs);

        Assert.Equal("STYLED", run.Text);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(OfficeTextDecorationStyle.Dotted, run.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, run.StrikeStyle);
        Assert.Equal(PdfTextBaseline.Superscript, run.Baseline);
        Assert.Equal(PdfColor.FromRgb(51, 102, 153), run.Color);
        Assert.Equal(PdfStandardFont.Helvetica, run.Font);
        Assert.Null(run.FontFamily);
        Assert.Equal(14D, run.FontSize);
    }
}
