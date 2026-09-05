using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfTextFormattingTests {
    [Fact]
    public void NativeRunFormattingProjectsToTypedPdfRuns() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(
            new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextRun authored = presentation.AddSlide()
            .AddTextBoxPoints("Styled", 10, 10, 200, 40)
            .Paragraphs[0].Runs[0];
        authored.Bold = true;
        authored.Italic = true;
        authored.UnderlineStyle = PowerPointUnderlineStyle.Wavy;
        authored.StrikeStyle = PowerPointStrikeStyle.Double;
        authored.SetSubscript();
        authored.Capitalization = PowerPointCapitalization.SmallCaps;
        authored.Color = "336699";
        authored.FontName = "Aptos";
        authored.FontSizePoints = 14D;

        PdfCore.PdfDocumentConversionResult conversion = presentation.ToPdfDocumentResult();
        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(conversion.Value.Blocks));
        PdfCore.PdfTextRun run = Assert.Single(
            canvas.Items.OfType<PdfCore.PdfCanvasTextBoxItem>().SelectMany(item => item.Runs),
            item => item.Text == "STYLED");

        Assert.Equal("STYLED", run.Text);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(OfficeTextDecorationStyle.Wavy, run.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, run.StrikeStyle);
        Assert.Equal(PdfCore.PdfTextBaseline.Subscript, run.Baseline);
        Assert.Equal(PdfCore.PdfColor.FromRgb(51, 102, 153), run.Color);
        Assert.Equal("Aptos", run.FontFamily);
        Assert.Equal(14D, run.FontSize);
        Assert.Contains(conversion.Warnings, warning => warning.Code == "small-caps-approximation");
    }

    [Fact]
    public void NativeRunCapitalizationUsesTheRunLanguage() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(
            new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextRun authored = presentation.AddSlide()
            .AddTextBoxPoints("i", 10, 10, 200, 40)
            .Paragraphs[0].Runs[0];
        authored.Capitalization = PowerPointCapitalization.AllCaps;
        authored.Language = "tr-TR";

        PdfCore.PdfDocumentConversionResult conversion = presentation.ToPdfDocumentResult();
        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(conversion.Value.Blocks));
        PdfCore.PdfTextRun run = Assert.Single(
            canvas.Items.OfType<PdfCore.PdfCanvasTextBoxItem>().SelectMany(item => item.Runs),
            item => item.Text == "İ");

        Assert.Equal("İ", run.Text);
    }

    [Fact]
    public void NoncanonicalBaselinePercentagesAreReportedForTextBoxesAndTables() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(
            new MemoryStream(), new PowerPointCreateOptions());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBoxPoints("Text", 10, 10, 120, 40).Paragraphs[0].Runs[0].BaselinePercent = 5D;
        PowerPointTableCell cell = slide.AddTablePoints(1, 1, 150, 10, 120, 40).GetCell(0, 0);
        cell.Text = "Cell";
        cell.Paragraphs[0].Runs[0].BaselinePercent = 80D;

        PdfCore.PdfDocumentConversionResult conversion = presentation.ToPdfDocumentResult();

        Assert.Single(conversion.Warnings, warning => warning.Code == "baseline-percent-approximation");
    }

    [Fact]
    public void TableRunFormattingResolvesParagraphListAndThemeDefaults() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(
            new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTableCell cell = presentation.AddSlide()
            .AddTablePoints(1, 1, 10, 10, 200, 60)
            .GetCell(0, 0);
        cell.Text = "i";
        A.Paragraph paragraph = Assert.Single(cell.Cell.TextBody!.Elements<A.Paragraph>());
        paragraph.ParagraphProperties = new A.ParagraphProperties(
            new A.DefaultRunProperties {
                Bold = true,
                Underline = A.TextUnderlineValues.Wavy,
                Capital = A.TextCapsValues.All
            });
        cell.Cell.TextBody.ListStyle!.PrependChild(new A.DefaultParagraphProperties(
            new A.DefaultRunProperties { Language = "tr-TR" }));
        cell.Cell.TextBody.ListStyle!.Append(new A.Level1ParagraphProperties(
            new A.DefaultRunProperties {
                Strike = A.TextStrikeValues.DoubleStrike,
                Baseline = 30000
            }));
        P.OtherStyle otherStyle = cell.SlidePart!.SlideLayoutPart!.SlideMasterPart!
            .SlideMaster!.TextStyles!.OtherStyle!;
        A.Level1ParagraphProperties themeLevel = otherStyle.GetFirstChild<A.Level1ParagraphProperties>()
            ?? otherStyle.AppendChild(new A.Level1ParagraphProperties());
        A.DefaultRunProperties themeDefaults = themeLevel.GetFirstChild<A.DefaultRunProperties>()
            ?? themeLevel.AppendChild(new A.DefaultRunProperties());
        themeDefaults.Italic = true;
        themeDefaults.FontSize = 1800;
        themeDefaults.Append(new A.LatinFont { Typeface = "+mn-lt" });
        themeDefaults.Append(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }));
        A.ThemeElements theme = cell.SlidePart.SlideLayoutPart.SlideMasterPart.ThemePart!.Theme!.ThemeElements!;
        theme.FontScheme!.MinorFont!.LatinFont!.Typeface = "Theme Minor";
        cell.SlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster!.ColorMap!.Accent1 = A.ColorSchemeIndexValues.Accent2;
        A.Accent2Color accent = theme.ColorScheme!.GetFirstChild<A.Accent2Color>()!;
        accent.RemoveAllChildren();
        accent.Append(new A.RgbColorModelHex { Val = "123456" });
        Assert.Single(paragraph.Elements<A.Run>()).RunProperties = new A.RunProperties();
        A.EndParagraphRunProperties? end = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
        var field = new A.Field(new A.RunProperties(), new A.Text("i")) {
            Id = "{11111111-1111-1111-1111-111111111111}",
            Type = "slidenum"
        };
        if (end != null) paragraph.InsertBefore(field, end);
        else paragraph.Append(field);

        PdfCore.PdfDocument document = presentation.ToPdfDocument();
        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(document.Blocks));
        PdfCore.PdfTextRun[] runs = Assert.Single(
            Assert.Single(canvas.Items.OfType<PdfCore.PdfCanvasTableItem>()).Block.Cells).Single().Runs.ToArray();

        Assert.Equal(new[] { "İ", "İ" }, runs.Select(run => run.Text));
        Assert.All(runs, run => {
            Assert.True(run.Bold);
            Assert.True(run.Italic);
            Assert.Equal(OfficeTextDecorationStyle.Wavy, run.UnderlineStyle);
            Assert.Equal(OfficeTextDecorationStyle.Double, run.StrikeStyle);
            Assert.Equal(PdfCore.PdfTextBaseline.Superscript, run.Baseline);
            Assert.Equal(18D, run.FontSize);
            Assert.Equal("Theme Minor", run.FontFamily);
            Assert.Equal(PdfCore.PdfColor.FromRgb(18, 52, 86), run.Color);
        });
    }

    [Fact]
    public void TextBoxRunsAndFieldsResolveParagraphListAndMasterDefaults() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(
            new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextBox textBox = presentation.AddSlide()
            .AddTextBoxPoints("i", 10, 10, 200, 60);
        A.Paragraph paragraph = Assert.Single(textBox.TextBody!.Elements<A.Paragraph>());
        paragraph.ParagraphProperties = new A.ParagraphProperties(
            new A.DefaultRunProperties {
                Bold = true,
                Underline = A.TextUnderlineValues.Wavy,
                Capital = A.TextCapsValues.All
            });
        textBox.TextBody.ListStyle!.PrependChild(new A.DefaultParagraphProperties(
            new A.DefaultRunProperties { Language = "tr-TR" }));
        textBox.TextBody.ListStyle.Append(new A.Level1ParagraphProperties(
            new A.DefaultRunProperties {
                Strike = A.TextStrikeValues.DoubleStrike,
                Baseline = 30000
            }));
        P.OtherStyle otherStyle = Assert.IsType<P.OtherStyle>(textBox.MasterTextStyle);
        A.Level1ParagraphProperties masterLevel = otherStyle.GetFirstChild<A.Level1ParagraphProperties>()
            ?? otherStyle.AppendChild(new A.Level1ParagraphProperties());
        A.DefaultRunProperties masterDefaults = masterLevel.GetFirstChild<A.DefaultRunProperties>()
            ?? masterLevel.AppendChild(new A.DefaultRunProperties());
        masterDefaults.Italic = true;
        masterDefaults.FontSize = 1800;
        masterDefaults.RemoveAllChildren<A.LatinFont>();
        masterDefaults.Append(new A.LatinFont { Typeface = "Aptos" });
        Assert.Single(paragraph.Elements<A.Run>()).RunProperties = new A.RunProperties();
        textBox.Paragraphs[0].AddField("i", "slidenum", "{33333333-3333-3333-3333-333333333333}");

        PdfCore.PdfDocument document = presentation.ToPdfDocument();
        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(document.Blocks));
        PdfCore.PdfTextRun[] runs = canvas.Items.OfType<PdfCore.PdfCanvasTextBoxItem>()
            .SelectMany(item => item.Runs)
            .Where(run => run.Text == "İ")
            .ToArray();

        Assert.Equal(2, runs.Length);
        Assert.All(runs, run => {
            Assert.True(run.Bold);
            Assert.True(run.Italic);
            Assert.Equal(OfficeTextDecorationStyle.Wavy, run.UnderlineStyle);
            Assert.Equal(OfficeTextDecorationStyle.Double, run.StrikeStyle);
            Assert.Equal(PdfCore.PdfTextBaseline.Superscript, run.Baseline);
            Assert.Equal(18D, run.FontSize);
            Assert.Equal("Aptos", run.FontFamily);
        });
    }

    [Fact]
    public void TableInheritedAndFieldFontsParticipateInPdfFontPreflight() {
        const string inheritedFamily = "OfficeIMO Missing Inherited Table Face";
        const string fieldFamily = "OfficeIMO Missing Field Face";
        using PowerPointPresentation presentation = PowerPointPresentation.Create(
            new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTableCell cell = presentation.AddSlide()
            .AddTablePoints(1, 1, 10, 10, 200, 60)
            .GetCell(0, 0);
        cell.Text = "Run";
        A.Paragraph paragraph = Assert.Single(cell.Cell.TextBody!.Elements<A.Paragraph>());
        P.OtherStyle otherStyle = cell.SlidePart!.SlideLayoutPart!.SlideMasterPart!
            .SlideMaster!.TextStyles!.OtherStyle!;
        A.Level1ParagraphProperties level = otherStyle.GetFirstChild<A.Level1ParagraphProperties>()
            ?? otherStyle.AppendChild(new A.Level1ParagraphProperties());
        A.DefaultRunProperties defaults = level.GetFirstChild<A.DefaultRunProperties>()
            ?? level.AppendChild(new A.DefaultRunProperties());
        defaults.RemoveAllChildren<A.LatinFont>();
        defaults.Append(new A.LatinFont { Typeface = inheritedFamily });
        var field = new A.Field(
            new A.RunProperties(new A.LatinFont { Typeface = fieldFamily }),
            new A.Text("Field")) {
            Id = "{22222222-2222-2222-2222-222222222222}",
            Type = "slidenum"
        };
        A.EndParagraphRunProperties? end = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
        if (end != null) paragraph.InsertBefore(field, end);
        else paragraph.Append(field);

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(
            new PowerPointToPdfOptions {
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
            });

        Assert.Contains(result.Warnings, warning => warning.Code == "font-family-substitution"
            && warning.Details.TryGetValue("fontFamily", out string? family) && family == inheritedFamily);
        Assert.Contains(result.Warnings, warning => warning.Code == "font-family-substitution"
            && warning.Details.TryGetValue("fontFamily", out string? family) && family == fieldFamily);
    }
}
