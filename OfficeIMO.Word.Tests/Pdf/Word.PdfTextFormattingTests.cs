using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
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

    [Fact]
    public void NativeCharacterAndTableStylesKeepTypedDecorationVariants() {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(
            new Style(
                new StyleName { Val = "Decorated character" },
                new StyleRunProperties(
                    new Underline { Val = UnderlineValues.Dash },
                    new DoubleStrike())) {
                Type = StyleValues.Character,
                StyleId = "DecoratedCharacter"
            },
            new Style(
                new StyleName { Val = "Decorated table" },
                new StyleRunProperties(
                    new Underline { Val = UnderlineValues.Wave },
                    new DoubleStrike())) {
                Type = StyleValues.Table,
                StyleId = "DecoratedTable"
            });

        WordParagraph characterRun = document.AddParagraph().AddText("Character style");
        characterRun.SetCharacterStyleId("DecoratedCharacter");
        MethodInfo createRuns = typeof(WordPdfConverterExtensions).GetMethod(
            "CreateNativeCellParagraphRuns",
            BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            new[] { typeof(WordParagraph), typeof(Dictionary<long, int>) },
            modifiers: null)!;
        PdfTextRun characterPdfRun = Assert.Single(Assert.IsAssignableFrom<IReadOnlyList<PdfTextRun>>(
            createRuns.Invoke(null, new object?[] { characterRun, null })));

        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "DecoratedTable" };
        Type converterType = typeof(WordPdfConverterExtensions);
        Type documentDefaultsType = converterType.GetNestedType("NativeDocumentDefaults", BindingFlags.NonPublic)!;
        object wordDefault = documentDefaultsType.GetProperty("WordDefault", BindingFlags.Public | BindingFlags.Static)!.GetValue(null)!;
        MethodInfo getTableDefaults = converterType.GetMethod(
            "GetNativeTableStyleDefaults",
            BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            new[] { typeof(WordTable), documentDefaultsType, typeof(bool) },
            modifiers: null)!;
        object tableDefaults = getTableDefaults.Invoke(null, new[] { table, wordDefault, false })!;
        object tableRunStyle = tableDefaults.GetType().GetProperty("RunStyle")!.GetValue(tableDefaults)!;

        Assert.Equal(OfficeTextDecorationStyle.Dashed, characterPdfRun.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, characterPdfRun.StrikeStyle);
        Assert.Equal(OfficeTextDecorationStyle.Wavy, tableRunStyle.GetType().GetProperty("UnderlineStyle")!.GetValue(tableRunStyle));
        Assert.Equal(OfficeTextDecorationStyle.Double, tableRunStyle.GetType().GetProperty("StrikeStyle")!.GetValue(tableRunStyle));
    }
}
