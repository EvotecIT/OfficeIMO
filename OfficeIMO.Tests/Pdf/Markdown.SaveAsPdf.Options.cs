using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class MarkdownSaveAsPdfOptionsTests {
    [Fact]
    public void ToPdfDocument_Markdown_ClonesCallerPdfOptionsBeforeApplyingAdapterDefaults() {
        var pdfOptions = new PdfCore.PdfOptions();
        var options = new MarkdownPdfSaveOptions {
            PdfOptions = pdfOptions,
            CreateOutlineFromHeadings = true
        };

        "# Heading".ToPdfDocument(options).ToBytes();

        Assert.False(pdfOptions.CreateOutlineFromHeadings);
    }

    [Fact]
    public void ToPdfDocument_Markdown_FontFamilyUsesSharedOfficeFontMapping() {
        var options = new MarkdownPdfSaveOptions {
            FontFamily = "Georgia",
            PdfOptions = new PdfCore.PdfOptions {
                CompressContentStreams = false
            }
        };

        byte[] bytes = "# Heading\n\nBody".ToPdfDocument(options).ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);

        Assert.True(
            raw.Contains("/BaseFont /Georgia-Regular", StringComparison.Ordinal) ||
            raw.Contains("/BaseFont /Times-Roman", StringComparison.Ordinal),
            "Expected Markdown font-family export to use an installed Georgia TrueType font or the mapped Times standard family.");
    }

    [Fact]
    public void ToPdfDocument_Markdown_DefaultsUseSharedUnicodeFallbackWhenAvailable() {
        var probe = new PdfCore.PdfOptions();
        if (!probe.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = ("# Faktura\n\n" + polish).ToPdfDocument(new MarkdownPdfSaveOptions()).ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains(polish, text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_HtmlTableCellFillsTrackBodyRowSpans() {
        const string html = """
<table>
  <tr><th>Group</th><th>Task</th></tr>
  <tr><td rowspan="2">A</td><td style="background:#ff0000">B</td></tr>
  <tr><td style="background:#0000ff">C</td></tr>
</table>
""";
        MarkdownDoc document = html.LoadFromHtml();
        byte[] bytes = document.ToPdfDocument(new MarkdownPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                CompressContentStreams = false,
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            }
        }).ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        var blueFills = ExtractFilledRectangles(raw, "0 0 1 rg");

        var blueFill = Assert.Single(blueFills);
        Assert.True(blueFill.X > 100D, "Expected the second-row blue cell fill to be painted in the second logical column.");
        Assert.True(blueFill.W > 40D);
    }

    [Fact]
    public void ToPdfDocument_HtmlTableCellAlignmentsTrackBodyRowSpans() {
        const string defaultHtml = """
<table>
  <colgroup><col style="width:80pt"><col style="width:130pt"><col style="width:60pt"></colgroup>
  <tr><th>Group</th><th>Task</th><th>Qty</th></tr>
  <tr><td rowspan="2">A</td><td>Build</td><td>1</td></tr>
  <tr><td>Support</td><td>2</td></tr>
</table>
""";
        const string alignedHtml = """
<table>
  <colgroup><col style="width:80pt"><col style="width:130pt"><col style="width:60pt"></colgroup>
  <tr><th>Group</th><th>Task</th><th>Qty</th></tr>
  <tr><td rowspan="2">A</td><td>Build</td><td>1</td></tr>
  <tr><td style="text-align:right">Support</td><td>2</td></tr>
</table>
""";

        var options = new MarkdownPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 240,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            }
        };

        byte[] defaultBytes = defaultHtml.LoadFromHtml().ToPdfDocument(options).ToBytes();
        byte[] alignedBytes = alignedHtml.LoadFromHtml().ToPdfDocument(options).ToBytes();

        using PdfPigDocument defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using PdfPigDocument alignedPdf = PdfPigDocument.Open(new MemoryStream(alignedBytes));

        double defaultX = FindWordStartX(defaultPdf.GetPage(1), "Support");
        double alignedX = FindWordStartX(alignedPdf.GetPage(1), "Support");

        Assert.True(alignedX > defaultX + 30D, $"Expected right-aligned Support cell to move within the second logical column. Default x: {defaultX:0.##}, aligned x: {alignedX:0.##}.");
    }

    [Fact]
    public void ToPdfDocument_HtmlPromotedHeaderRowSpanDoesNotCrossPdfHeaderBoundary() {
        const string html = """
<table>
  <tr><td rowspan="2">Group</td><td>Task</td></tr>
  <tr><td>Setup</td></tr>
</table>
""";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Equal(1, table.HeaderCells[0].RowSpan);
        byte[] bytes = document.ToPdfDocument(new MarkdownPdfSaveOptions()).ToBytes();
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(bytes));
    }

    private static IReadOnlyList<(double X, double Y, double W, double H)> ExtractFilledRectangles(string rawPdf, string colorOperator) {
        var rectangles = new List<(double X, double Y, double W, double H)>();
        string pattern = Regex.Escape(colorOperator) +
            @"\s+(?<x>-?\d+(?:\.\d+)?)\s+(?<y>-?\d+(?:\.\d+)?)\s+(?<w>-?\d+(?:\.\d+)?)\s+(?<h>-?\d+(?:\.\d+)?)\s+re\s+f";

        foreach (Match match in Regex.Matches(rawPdf, pattern, RegexOptions.Singleline)) {
            rectangles.Add((
                ParseInvariantDouble(match.Groups["x"].Value),
                ParseInvariantDouble(match.Groups["y"].Value),
                ParseInvariantDouble(match.Groups["w"].Value),
                ParseInvariantDouble(match.Groups["h"].Value)));
        }

        return rectangles;
    }

    private static double ParseInvariantDouble(string value) =>
        double.Parse(value, CultureInfo.InvariantCulture);

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }
}
