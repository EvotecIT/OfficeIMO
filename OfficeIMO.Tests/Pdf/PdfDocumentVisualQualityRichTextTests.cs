using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void RichParagraph_ResetColor_ReturnsToDefaultTextColor() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(p => p
                .Text("Before ")
                .Color(new PdfColor(1, 0, 0))
                .Text("Red")
                .ResetColor()
                .Text("After"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int redText = content.IndexOf("<526564>", StringComparison.Ordinal);
        int afterText = content.IndexOf("<4166746572>", StringComparison.Ordinal);

        Assert.True(redText >= 0, "Expected encoded 'Red' text in the generated PDF content stream.");
        Assert.True(afterText > redText, "Expected encoded 'After' text after the red run.");

        int redColorBeforeRed = content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal);
        int blackColorBeforeAfter = content.LastIndexOf("0 0 0 rg", afterText, StringComparison.Ordinal);
        int redColorBeforeAfter = content.LastIndexOf("1 0 0 rg", afterText, StringComparison.Ordinal);

        Assert.True(redColorBeforeRed >= 0, "Expected the red run to emit a red fill color.");
        Assert.True(blackColorBeforeAfter > redText, "Expected ResetColor to emit black/default fill color before the following run.");
        Assert.True(redColorBeforeAfter < blackColorBeforeAfter, "Expected the following run not to inherit the previous red fill color.");
    }

    [Fact]
    public void RichParagraph_FontSize_AppliesOnlyToScopedRuns() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFontSize = 11
            })
            .Paragraph(p => p
                .Text("Small")
                .FontSize(18)
                .Text("Large")
                .ResetFontSize()
                .Text("Normal"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int smallText = content.IndexOf("<536D616C6C>", StringComparison.Ordinal);
        int largeText = content.IndexOf("<4C61726765>", StringComparison.Ordinal);
        int normalText = content.IndexOf("<4E6F726D616C>", StringComparison.Ordinal);

        Assert.True(smallText >= 0, "Expected encoded 'Small' text in the generated PDF content stream.");
        Assert.True(largeText > smallText, "Expected encoded 'Large' text after the default-sized run.");
        Assert.True(normalText > largeText, "Expected encoded 'Normal' text after the large run.");

        int defaultSizeBeforeSmall = content.LastIndexOf("/F1 11 Tf", smallText, StringComparison.Ordinal);
        int largeSizeBeforeLarge = content.LastIndexOf("/F1 18 Tf", largeText, StringComparison.Ordinal);
        int defaultSizeBeforeNormal = content.LastIndexOf("/F1 11 Tf", normalText, StringComparison.Ordinal);

        Assert.True(defaultSizeBeforeSmall >= 0, "Expected the first run to use the paragraph/default font size.");
        Assert.True(largeSizeBeforeLarge > smallText, "Expected FontSize(18) to emit an 18-point font before the scoped run.");
        Assert.True(defaultSizeBeforeNormal > largeText, "Expected ResetFontSize to restore the paragraph/default font size for later runs.");
    }

    [Fact]
    public void RichParagraph_BackgroundColor_RendersBehindScopedRuns() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(p => p
                .Text("Before ")
                .BackgroundColor(PdfColor.FromRgb(255, 255, 0))
                .Text("Marked")
                .ResetBackgroundColor()
                .Text("After"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int markedText = content.IndexOf("<4D61726B6564>", StringComparison.Ordinal);
        int afterText = content.IndexOf("<4166746572>", StringComparison.Ordinal);

        Assert.True(markedText >= 0, "Expected encoded 'Marked' text in the generated PDF content stream.");
        Assert.True(afterText > markedText, "Expected encoded 'After' text after the highlighted run.");

        int highlightFill = content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal);
        int highlightRect = content.LastIndexOf(" re f", markedText, StringComparison.Ordinal);

        Assert.True(highlightFill >= 0, "Expected the highlighted run to emit a yellow fill color.");
        Assert.True(highlightRect > highlightFill, "Expected the highlighted run to emit a filled rectangle before the text.");
        Assert.Single(Regex.Matches(content, "1 1 0 rg").Cast<Match>());
    }

    [Fact]
    public void RichParagraph_BackgroundColor_MergesMultiWordRunsIntoContinuousHighlight() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(p => p
                .Text("Before ")
                .BackgroundColor(PdfColor.FromRgb(255, 255, 0))
                .Text("Marked Words")
                .ResetBackgroundColor()
                .Text(" After"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int markedText = content.IndexOf("<4D61726B6564>", StringComparison.Ordinal);
        int wordsText = content.IndexOf("<576F726473>", StringComparison.Ordinal);

        Assert.True(markedText >= 0, "Expected encoded 'Marked' text in the generated PDF content stream.");
        Assert.True(wordsText > markedText, "Expected encoded 'Words' text after the first highlighted word.");
        Assert.Single(Regex.Matches(content, "1 1 0 rg").Cast<Match>());
        Match highlightRect = Regex.Match(content, @"1 1 0 rg\s+([0-9.]+) ([0-9.]+) ([0-9.]+) ([0-9.]+) re\s+f");
        Assert.True(highlightRect.Success, "Expected one continuous yellow rectangle for the whole highlighted phrase.");
        double width = double.Parse(highlightRect.Groups[3].Value, CultureInfo.InvariantCulture);
        Assert.True(width > 76D, $"Expected the highlight rectangle to include the space between highlighted words. Width: {width:0.##}.");
    }

    [Fact]
    public void Table_RichTextCell_RendersScopedRunStyles() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFontSize = 11
            })
            .Table(new[] {
                new[] {
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal("Plain "),
                        new TextRun("CellRed", color: PdfColor.FromRgb(255, 0, 0)),
                        TextRun.Normal(" "),
                        TextRun.Bolded("CellBold"),
                        TextRun.Normal(" "),
                        TextRun.Normal("CellMarked", backgroundColor: PdfColor.FromRgb(255, 255, 0)),
                        TextRun.Normal(" "),
                        TextRun.Normal("CellLarge", fontSize: 18)
                    })
                }
            }, style: new PdfTableStyle {
                HeaderRowCount = 0
            })
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int redText = content.IndexOf("<43656C6C526564>", StringComparison.Ordinal);
        int boldText = content.IndexOf("<43656C6C426F6C64>", StringComparison.Ordinal);
        int markedText = content.IndexOf("<43656C6C4D61726B6564>", StringComparison.Ordinal);
        int largeText = content.IndexOf("<43656C6C4C61726765>", StringComparison.Ordinal);

        Assert.True(redText >= 0, "Expected encoded rich table cell red text in the PDF content stream.");
        Assert.True(boldText > redText, "Expected encoded bold table cell text after the red run.");
        Assert.True(markedText > boldText, "Expected encoded highlighted table cell text after the bold run.");
        Assert.True(largeText > markedText, "Expected encoded large table cell text after the highlighted run.");

        Assert.True(content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal) >= 0, "Expected rich table cell color to emit a red fill color.");
        Assert.True(content.LastIndexOf("/F2 11 Tf", boldText, StringComparison.Ordinal) >= 0, "Expected rich table cell bold run to use the bold font resource.");
        Assert.True(content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal) >= 0, "Expected rich table cell highlight to emit a yellow fill color.");
        Assert.True(content.LastIndexOf("/F1 18 Tf", largeText, StringComparison.Ordinal) >= 0, "Expected rich table cell font size to emit an 18-point run.");
    }

    [Fact]
    public void RichText_RejectsNullRunTextBeforeRendering() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Paragraph(p => p.Text(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Paragraph(p => p.Bold(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Paragraph(p => p.Italic(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Paragraph(p => p.Underlined(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Paragraph(p => p.Strikethrough(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            TextRun.Normal(null!));
    }

    [Fact]
    public void RichText_RejectsInvalidBaselineValuesBeforeRendering() {
        var exception = Assert.Throws<ArgumentException>(() =>
            new TextRun("Invalid baseline", baseline: (PdfTextBaseline)99));

        Assert.Equal("baseline", exception.ParamName);
        Assert.Contains("PDF text baseline must be Normal, Superscript, or Subscript.", exception.Message, StringComparison.Ordinal);

        var builderException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.Baseline((PdfTextBaseline)99).Text("Invalid baseline")));

        Assert.Equal("baseline", builderException.ParamName);
        Assert.Contains("PDF text baseline must be Normal, Superscript, or Subscript.", builderException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RichText_RendersSuperscriptAndSubscriptWithTextRise() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFontSize = 14
            })
            .Paragraph(p => p
                .Text("Base")
                .Superscript()
                .Link("SUP", "https://example.com/sup")
                .Superscript(false)
                .Text("Mid")
                .Subscript("SUB")
                .Text("End"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/URI (https://example.com/sup)", content);
        Assert.Contains("4.9 Ts", content);
        Assert.Contains("-2.52 Ts", content);
        Assert.Contains("0 Ts", content);
        Assert.Contains("/F1 9.1 Tf", content);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();

        double baseY = AverageBaselineY(letters, "Base");
        double superY = AverageBaselineY(letters, "SUP");
        double midY = AverageBaselineY(letters, "Mid");
        double subY = AverageBaselineY(letters, "SUB");

        Assert.InRange(Math.Abs(baseY - midY), 0, 0.75);
        Assert.True(superY > baseY + 3.5, $"Expected superscript text to rise above the normal baseline. Base: {baseY:0.##}, super: {superY:0.##}.");
        Assert.True(subY < baseY - 1.5, $"Expected subscript text to sit below the normal baseline. Base: {baseY:0.##}, sub: {subY:0.##}.");
    }

    [Fact]
    public void ParagraphBlocks_SnapshotRunsIntoReadOnlyCollections() {
        var runs = new List<TextRun> {
            TextRun.Normal("Stable alpha"),
            TextRun.Bolded("Stable beta")
        };
        var paragraphStyle = new PdfParagraphStyle {
            LineHeight = 1.6,
            LeftIndent = 4,
            RightIndent = 5,
            FirstLineIndent = 3,
            SpacingBefore = 6,
            SpacingAfter = 7,
            KeepTogether = true,
            KeepWithNext = true,
            WidowControl = true
        };
        var panelStyle = new PanelStyle();

        var paragraph = new RichParagraphBlock(runs, PdfAlign.Left, null, paragraphStyle);
        var panel = new PanelParagraphBlock(runs, PdfAlign.Left, null, panelStyle);

        runs[0] = TextRun.Normal("Mutated alpha");
        runs.Add(TextRun.Normal("Late gamma"));
        paragraphStyle.LineHeight = 2.2;
        paragraphStyle.LeftIndent = 20;
        paragraphStyle.RightIndent = 21;
        paragraphStyle.FirstLineIndent = 22;
        paragraphStyle.SpacingBefore = 22;
        paragraphStyle.SpacingAfter = 23;
        paragraphStyle.KeepTogether = false;
        paragraphStyle.KeepWithNext = false;
        paragraphStyle.WidowControl = false;

        Assert.Equal(new[] { "Stable alpha", "Stable beta" }, paragraph.Runs.Select(run => run.Text).ToArray());
        Assert.Equal(new[] { "Stable alpha", "Stable beta" }, panel.Runs.Select(run => run.Text).ToArray());
        Assert.False(paragraph.Runs is List<TextRun>);
        Assert.False(panel.Runs is List<TextRun>);
        Assert.Equal(1.6, paragraph.Style!.LineHeight);
        Assert.Equal(4, paragraph.Style.LeftIndent);
        Assert.Equal(5, paragraph.Style.RightIndent);
        Assert.Equal(3, paragraph.Style.FirstLineIndent);
        Assert.Equal(6, paragraph.Style.SpacingBefore);
        Assert.Equal(7, paragraph.Style.SpacingAfter);
        Assert.True(paragraph.Style.KeepTogether);
        Assert.True(paragraph.Style.KeepWithNext);
        Assert.True(paragraph.Style.WidowControl);
    }


}
