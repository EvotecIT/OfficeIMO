using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Html;
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
        var pdfOptions = new PdfCore.PdfCreateOptions();
        var options = new MarkdownPdfExportOptions {
            PdfOptions = pdfOptions,
            CreateOutlineFromHeadings = true
        };

        OfficeIMO.Markdown.MarkdownDocument.Parse("# Heading").ToPdf(options).ToBytes();

        Assert.False(pdfOptions.CreateOutlineFromHeadings);
    }

    [Fact]
    public void ToPdfDocument_Markdown_FontFamilyUsesSharedOfficeFontMapping() {
        var options = new MarkdownPdfExportOptions {
            FontFamily = "Georgia",
            PdfOptions = new PdfCore.PdfCreateOptions {
                CompressContentStreams = false
            }
        };

        using PdfCore.PdfDocument document = OfficeIMO.Markdown.MarkdownDocument.Parse("# Heading\n\nBody").ToPdf(options);
        // Georgia selects the logical serif family. The embedded face can be a
        // theme serif fallback when Georgia is not installed on the host.
        Assert.Equal(PdfCore.PdfStandardFont.TimesRoman, document.Options.DefaultFont);
        using var rendered = PdfPigDocument.Open(document.ToBytes());
        string text = rendered.GetPage(1).Text;
        Assert.Contains("Heading", text, StringComparison.Ordinal);
        Assert.Contains("Body", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_Markdown_PortableDeterministicResourcePolicyDoesNotEmbedSystemFonts() {
        var options = new MarkdownPdfExportOptions {
            FontFamily = "Georgia",
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };

        PdfCore.PdfDocument document = OfficeIMO.Markdown.MarkdownDocument.Parse("# Heading\n\nBody").ToPdf(options);

        Assert.Equal(PdfCore.PdfStandardFont.TimesRoman, document.Options.DefaultFont);
        Assert.False(document.Options.HasEmbeddedStandardFontFamily(PdfCore.PdfStandardFont.TimesRoman));
    }

    [Fact]
    public void ToPdfDocument_Markdown_TextFallbacksPreserveCallerPdfOptionsFontSlots() {
        var options = new MarkdownPdfExportOptions {
            Style = MarkdownPdfStyle.Plain(),
            PdfOptions = new PdfCore.PdfCreateOptions {
                DefaultFont = PdfCore.PdfStandardFont.TimesRoman,
                HeaderFont = PdfCore.PdfStandardFont.Courier,
                FooterFont = PdfCore.PdfStandardFont.Helvetica
            }
        };

        PdfCore.PdfDocument document = OfficeIMO.Markdown.MarkdownDocument.Parse("# Heading\n\nBody").ToPdf(options);

        Assert.Equal(PdfCore.PdfStandardFont.TimesRoman, document.Options.DefaultFont);
        Assert.Equal(PdfCore.PdfStandardFont.Courier, document.Options.HeaderFont);
        Assert.Equal(PdfCore.PdfStandardFont.Helvetica, document.Options.FooterFont);
    }

    [Fact]
    public void ToPdfDocument_Markdown_TextFallbacksReserveCourierForCodeText() {
        PdfCore.PdfDocument document = OfficeIMO.Markdown.MarkdownDocument.Parse("`code` text").ToPdf(new MarkdownPdfExportOptions());
        PdfCore.PdfEmbeddedFontFallbackSet? fallbackSet = document.Options.EmbeddedFontFallbacks;
        if (fallbackSet == null) {
            return;
        }

        Assert.DoesNotContain(
            PdfCore.PdfStandardFont.Courier,
            fallbackSet.FontSlots.Select(PdfCore.PdfStandardFontMapper.GetFontFamily));
    }

    [Fact]
    public void ToPdfDocument_Markdown_TextFallbacksEmbedMonospaceBeforeReservingCourierForSymbols() {
        var probe = new PdfCore.PdfCreateOptions();
        if (!probe.TryRegisterDefaultDocumentMonospaceFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        PdfCore.PdfDocument document = OfficeIMO.Markdown.MarkdownDocument.Parse("`Zażółć` text").ToPdf(new MarkdownPdfExportOptions());

        Assert.True(document.Options.HasEmbeddedStandardFontFamily(PdfCore.PdfStandardFont.Courier));
    }

    [Fact]
    public void ToPdfDocument_Markdown_DefaultsUseSharedUnicodeFallbackWhenAvailable() {
        var probe = new PdfCore.PdfCreateOptions();
        if (!probe.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = OfficeIMO.Markdown.MarkdownDocument.Parse("# Faktura\n\n" + polish).ToPdf(new MarkdownPdfExportOptions()).ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

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
        MarkdownDocument document = HtmlConversionDocument.Parse(html).ToMarkdown();
        byte[] bytes = document.ToPdf(new MarkdownPdfExportOptions {
            PdfOptions = new PdfCore.PdfCreateOptions {
                CompressContentStreams = false,
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
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

        var options = new MarkdownPdfExportOptions {
            PdfOptions = new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(240),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
            }
        };

        byte[] defaultBytes = HtmlConversionDocument.Parse(defaultHtml).ToMarkdown().ToPdf(options).ToBytes();
        byte[] alignedBytes = HtmlConversionDocument.Parse(alignedHtml).ToMarkdown().ToPdf(options).ToBytes();

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

        MarkdownDocument document = HtmlConversionDocument.Parse(html).ToMarkdown();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Equal(1, table.HeaderCells[0].RowSpan);
        byte[] bytes = document.ToPdf(new MarkdownPdfExportOptions()).ToBytes();
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(bytes));
    }

    [Fact]
    public void ToMarkdownDocument_TableColSpansRespectExpandedColumnLimit() {
        var html = new StringBuilder("<table><tr>");
        for (int i = 0; i < 20; i++) {
            html.Append("<td colspan=\"512\">x</td>");
        }
        html.Append("</tr></table>");

        MarkdownDocument document = HtmlConversionDocument.Parse(html.ToString()).ToMarkdown(new MarkdownHtmlImportOptions {
            MaxTableExpandedColumns = 32
        });

        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        Assert.True(table.SkippedColumnCount > 0);
        Assert.Equal(32, table.HeaderCells.Count);
        Assert.Equal(32, table.HeaderCells[0].ColumnSpan);

        byte[] bytes = document.ToPdf(new MarkdownPdfExportOptions()).ToBytes();
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(bytes));
    }

    [Fact]
    public void TableBlock_StructuredColSpansUseBoundedLogicalColumnCount() {
        var table = new TableBlock();
        var headerCells = new List<TableCell>();
        for (int i = 0; i < 1000; i++) {
            table.Headers.Add("x");
            headerCells.Add(new TableCell { ColumnSpan = 512 });
        }

        table.SetStructuredCells(headerCells, rows: null, table.ComputeContentSignature());

        Assert.Equal(TableBlock.MaxEffectiveColumnCount, table.HeaderCells.Count);
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
