using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfHtmlConverterTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Lists_And_Tables() {
        const string html = "<ul><li>Allergy</li><li><strong>Medication</strong></li></ul><table><tr><th>Name</th><th>Value</th></tr><tr><td>Pulse</td><td>72</td></tr></table>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Equal(RtfListKind.Bullet, document.Paragraphs[0].ListKind);
        Assert.Equal("Allergy", document.Paragraphs[0].ToPlainText());
        Assert.Equal(RtfListKind.Bullet, document.Paragraphs[1].ListKind);
        Assert.Contains(document.Paragraphs[1].Runs, run => run.Text == "Medication" && run.Bold);

        RtfTable table = Assert.IsType<RtfTable>(document.Blocks[2]);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("Name", table.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
        Assert.Equal("72", table.Rows[1].Cells[1].Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Header_And_Cell_Styles() {
        const string html = "<table><thead><tr><th style=\"background-color:#f2f2f2;width:25%;vertical-align:middle\">Name</th><th style=\"text-align:right;width:72pt\">Value</th></tr></thead><tbody><tr><td style=\"background:#fff2cc;vertical-align:bottom\">Pulse</td><td>72</td></tr></tbody></table>";

        RtfDocument document = html.ToRtfDocument();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        Assert.True(table.Rows[0].RepeatHeader);
        RtfTableCell firstHeader = table.Rows[0].Cells[0];
        Assert.Equal(1, firstHeader.BackgroundColorIndex);
        Assert.Equal(1250, firstHeader.PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Percent, firstHeader.PreferredWidthUnit);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, firstHeader.VerticalAlignment);
        Assert.Equal(RtfTextAlignment.Center, firstHeader.Paragraphs[0].Alignment);
        Assert.Contains(firstHeader.Paragraphs[0].Runs, run => run.Text == "Name" && run.Bold);

        RtfTableCell secondHeader = table.Rows[0].Cells[1];
        Assert.Equal(1440, secondHeader.PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Twips, secondHeader.PreferredWidthUnit);
        Assert.Equal(RtfTextAlignment.Right, secondHeader.Paragraphs[0].Alignment);

        RtfTableCell pulseCell = table.Rows[1].Cells[0];
        Assert.Equal(2, pulseCell.BackgroundColorIndex);
        Assert.Equal(RtfTableCellVerticalAlignment.Bottom, pulseCell.VerticalAlignment);

        string rtf = document.ToRtf();
        Assert.Contains(@"\trhdr", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth2\clwWidth1250", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth3\clwWidth1440", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvertalc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvertalb", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.True(roundTripTable.Rows[0].RepeatHeader);
        Assert.Equal(RtfTableWidthUnit.Percent, roundTripTable.Rows[0].Cells[0].PreferredWidthUnit);
        Assert.Equal(1250, roundTripTable.Rows[0].Cells[0].PreferredWidth);
        Assert.Equal(RtfTableCellVerticalAlignment.Bottom, roundTripTable.Rows[1].Cells[0].VerticalAlignment);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Header_And_Cell_Styles() {
        RtfDocument document = RtfDocument.Create();
        int headerBackground = document.AddColor(242, 242, 242);
        int bodyBackground = document.AddColor(255, 242, 204);
        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0]
            .SetBackgroundColor(headerBackground)
            .SetPreferredWidth(1250, RtfTableWidthUnit.Percent);
        table.Rows[0].Cells[0].VerticalAlignment = RtfTableCellVerticalAlignment.Center;
        table.Rows[0].Cells[0].AddParagraph("Name");
        table.Rows[0].Cells[1]
            .SetPreferredWidth(1440, RtfTableWidthUnit.Twips)
            .AddParagraph("Value");
        table.Rows[1].Cells[0].SetBackgroundColor(bodyBackground);
        table.Rows[1].Cells[0].VerticalAlignment = RtfTableCellVerticalAlignment.Bottom;
        table.Rows[1].Cells[0].AddParagraph("Pulse");
        table.Rows[1].Cells[1].AddParagraph("72");

        string html = document.ToHtml();

        Assert.Equal("<table><thead><tr><th style=\"background-color:#F2F2F2;width:25%;vertical-align:middle;\"><p>Name</p></th><th style=\"width:72pt;\"><p>Value</p></th></tr></thead><tbody><tr><td style=\"background-color:#FFF2CC;vertical-align:bottom;\"><p>Pulse</p></td><td><p>72</p></td></tr></tbody></table>", html);
    }

    [Fact]
    public void RtfDocument_ToHtml_Reports_Diagnostic_When_Image_Embedding_Is_Disabled() {
        RtfDocument document = RtfDocument.Create();
        document.AddImage(RtfImageFormat.Png, new byte[] { 0x89, 0x50, 0x4E, 0x47 });
        var callbackDiagnostics = new List<HtmlRtfConversionDiagnostic>();
        var options = new RtfToHtmlOptions {
            EmbedImagesAsDataUri = false,
            DiagnosticHandler = callbackDiagnostics.Add
        };

        string html = document.ToHtml(options);

        Assert.DoesNotContain("<img", html, StringComparison.Ordinal);
        HtmlRtfConversionDiagnostic diagnostic = Assert.Single(options.Diagnostics);
        Assert.Same(diagnostic, Assert.Single(callbackDiagnostics));
        Assert.Equal("RtfHtmlImageEmbeddingDisabled", diagnostic.Code);
        Assert.Equal(HtmlRtfConversionDiagnosticSeverity.Warning, diagnostic.Severity);
        Assert.Equal(RtfImageFormat.Png.ToString(), diagnostic.Source);
    }

    [Fact]
    public void RtfDocument_ToHtml_Reports_Diagnostic_For_Unsupported_Image_Format() {
        RtfDocument document = RtfDocument.Create();
        document.AddImage(RtfImageFormat.Emf, new byte[] { 0x01, 0x02, 0x03 });
        var options = new RtfToHtmlOptions();

        string html = document.ToHtml(options);

        Assert.DoesNotContain("<img", html, StringComparison.Ordinal);
        HtmlRtfConversionDiagnostic diagnostic = Assert.Single(options.Diagnostics);
        Assert.Equal("RtfHtmlImageFormatUnsupported", diagnostic.Code);
        Assert.Equal(RtfImageFormat.Emf.ToString(), diagnostic.Source);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Colspan_And_Rowspan() {
        const string html = "<table><tr><th colspan=\"2\">Panel</th><th rowspan=\"2\">Flag</th></tr><tr><td>Pulse</td><td>72</td></tr></table>";

        RtfDocument document = html.ToRtfDocument();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        Assert.Equal(RtfTableCellMerge.First, table.Rows[0].Cells[0].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, table.Rows[0].Cells[1].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.First, table.Rows[0].Cells[2].VerticalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, table.Rows[1].Cells[2].VerticalMerge);
        Assert.Equal("Panel", table.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
        Assert.Equal("Flag", table.Rows[0].Cells[2].Paragraphs[0].ToPlainText());

        string rtf = document.ToRtf();
        Assert.Contains(@"\clmgf", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clmrg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvmgf", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvmrg", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.Equal(RtfTableCellMerge.First, roundTripTable.Rows[0].Cells[0].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, roundTripTable.Rows[0].Cells[1].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.First, roundTripTable.Rows[0].Cells[2].VerticalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, roundTripTable.Rows[1].Cells[2].VerticalMerge);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Colspan_And_Rowspan() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 3);
        table.Rows[0].Cells[0].HorizontalMerge = RtfTableCellMerge.First;
        table.Rows[0].Cells[0].AddParagraph("Panel");
        table.Rows[0].Cells[1].HorizontalMerge = RtfTableCellMerge.Continue;
        table.Rows[0].Cells[2].VerticalMerge = RtfTableCellMerge.First;
        table.Rows[0].Cells[2].AddParagraph("Flag");
        table.Rows[1].Cells[0].AddParagraph("Pulse");
        table.Rows[1].Cells[1].AddParagraph("72");
        table.Rows[1].Cells[2].VerticalMerge = RtfTableCellMerge.Continue;

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr><td colspan=\"2\"><p>Panel</p></td><td rowspan=\"2\"><p>Flag</p></td></tr><tr><td><p>Pulse</p></td><td><p>72</p></td></tr></tbody></table>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Cell_Borders_And_Padding() {
        const string html = "<table><tr><td style=\"padding:6pt 9pt 3pt 12pt;border:1pt solid #0c2238;border-bottom:2pt dashed red\">Value</td></tr></table>";

        RtfDocument document = html.ToRtfDocument();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        RtfTableCell cell = table.Rows[0].Cells[0];
        Assert.Equal(120, cell.PaddingTopTwips);
        Assert.Equal(240, cell.PaddingLeftTwips);
        Assert.Equal(60, cell.PaddingBottomTwips);
        Assert.Equal(180, cell.PaddingRightTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, cell.TopBorder.Style);
        Assert.Equal(20, cell.TopBorder.Width);
        Assert.Equal(1, cell.TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, cell.BottomBorder.Style);
        Assert.Equal(40, cell.BottomBorder.Width);
        Assert.Equal(2, cell.BottomBorder.ColorIndex);

        string rtf = document.ToRtf();
        Assert.Contains(@"\clpadt120\clpadft3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadl240\clpadfl3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrt\brdrs\brdrw20\brdrcf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrb\brdrdash\brdrw40\brdrcf2", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        RtfTableCell readCell = roundTripTable.Rows[0].Cells[0];
        Assert.Equal(120, readCell.PaddingTopTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, readCell.TopBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, readCell.BottomBorder.Style);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Cell_Borders_And_Padding() {
        RtfDocument document = RtfDocument.Create();
        int dark = document.AddColor(12, 34, 56);
        int red = document.AddColor(255, 0, 0);
        RtfTable table = document.AddTable(1, 1);
        RtfTableCell cell = table.Rows[0].Cells[0];
        cell.SetPadding(topTwips: 120, leftTwips: 240, bottomTwips: 60, rightTwips: 180);
        cell.TopBorder.Style = RtfTableCellBorderStyle.Single;
        cell.TopBorder.Width = 20;
        cell.TopBorder.ColorIndex = dark;
        cell.LeftBorder.Style = RtfTableCellBorderStyle.Single;
        cell.LeftBorder.Width = 20;
        cell.LeftBorder.ColorIndex = dark;
        cell.BottomBorder.Style = RtfTableCellBorderStyle.Dashed;
        cell.BottomBorder.Width = 40;
        cell.BottomBorder.ColorIndex = red;
        cell.RightBorder.Style = RtfTableCellBorderStyle.Double;
        cell.RightBorder.Width = 20;
        cell.RightBorder.ColorIndex = dark;
        cell.AddParagraph("Value");

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr><td style=\"padding-top:6pt;padding-left:12pt;padding-bottom:3pt;padding-right:9pt;border-top:1pt solid #0C2238;border-left:1pt solid #0C2238;border-bottom:2pt dashed #FF0000;border-right:1pt double #0C2238;\"><p>Value</p></td></tr></tbody></table>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Legacy_Table_Cell_Attributes() {
        const string html = "<table><tr><td align=\"right\" valign=\"middle\" bgcolor=\"#fff2cc\" width=\"30%\" nowrap>Result</td></tr></table>";

        RtfDocument document = html.ToRtfDocument();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        RtfTableCell cell = table.Rows[0].Cells[0];
        Assert.Equal(RtfTextAlignment.Right, cell.Paragraphs[0].Alignment);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, cell.VerticalAlignment);
        Assert.Equal(1, cell.BackgroundColorIndex);
        Assert.Equal(1500, cell.PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Percent, cell.PreferredWidthUnit);
        Assert.True(cell.NoWrap);

        string rtf = document.ToRtf();
        Assert.Contains(@"\clftsWidth2\clwWidth1500", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clNoWrap", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvertalc", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.True(roundTripTable.Rows[0].Cells[0].NoWrap);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, roundTripTable.Rows[0].Cells[0].VerticalAlignment);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Cell_Nowrap_Style() {
        RtfDocument document = RtfDocument.Create();
        int background = document.AddColor(255, 242, 204);
        RtfTable table = document.AddTable(1, 1);
        RtfTableCell cell = table.Rows[0].Cells[0];
        cell.SetBackgroundColor(background)
            .SetPreferredWidth(1500, RtfTableWidthUnit.Percent)
            .SetNoWrap();
        cell.VerticalAlignment = RtfTableCellVerticalAlignment.Center;
        cell.AddParagraph("Result");

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr><td style=\"background-color:#FFF2CC;width:30%;vertical-align:middle;white-space:nowrap;--officeimo-rtf-cell-nowrap:true;\"><p>Result</p></td></tr></tbody></table>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Row_Attributes_And_Styles() {
        const string html = "<table><tr align=\"center\" bgcolor=\"#f2f2f2\" width=\"80%\" height=\"24pt\" style=\"padding:3pt 4pt 5pt 6pt\"><td>Result</td></tr></table>";

        RtfDocument document = html.ToRtfDocument();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        RtfTableRow row = table.Rows[0];
        Assert.Equal(RtfTableAlignment.Center, row.Alignment);
        Assert.Equal(1, row.BackgroundColorIndex);
        Assert.Equal(4000, row.PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Percent, row.PreferredWidthUnit);
        Assert.Equal(480, row.HeightTwips);
        Assert.Equal(60, row.PaddingTopTwips);
        Assert.Equal(120, row.PaddingLeftTwips);
        Assert.Equal(100, row.PaddingBottomTwips);
        Assert.Equal(80, row.PaddingRightTwips);

        string rtf = document.ToRtf();
        Assert.Contains(@"\trrh480", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trftsWidth2\trwWidth4000", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trcbpat1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trpaddt60\trpaddft3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trqc", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.Equal(RtfTableAlignment.Center, roundTripTable.Rows[0].Alignment);
        Assert.Equal(480, roundTripTable.Rows[0].HeightTwips);
        Assert.Equal(4000, roundTripTable.Rows[0].PreferredWidth);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Row_Styles() {
        RtfDocument document = RtfDocument.Create();
        int background = document.AddColor(242, 242, 242);
        RtfTable table = document.AddTable(1, 1);
        RtfTableRow row = table.Rows[0];
        row.SetAlignment(RtfTableAlignment.Right)
            .SetBackgroundColor(background)
            .SetPadding(topTwips: 60, leftTwips: 120, bottomTwips: 100, rightTwips: 80);
        row.PreferredWidth = 4000;
        row.PreferredWidthUnit = RtfTableWidthUnit.Percent;
        row.HeightTwips = 480;
        row.Cells[0].AddParagraph("Result");

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr style=\"background-color:#F2F2F2;text-align:right;width:80%;height:24pt;padding-top:3pt;padding-left:6pt;padding-bottom:5pt;padding-right:4pt;\"><td><p>Result</p></td></tr></tbody></table>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Image_Dimensions() {
        const string html = "<p><img src=\"data:image/png;base64,iVBORw==\" alt=\"Chart\" width=\"96\" height=\"48\" style=\"width:120pt;height:60pt\"></p>";

        RtfDocument document = html.ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfImage image = Assert.Single(paragraph.Inlines.OfType<RtfImage>());
        Assert.Equal(RtfImageFormat.Png, image.Format);
        Assert.Equal("Chart", image.Description);
        Assert.Equal(96, image.SourceWidth);
        Assert.Equal(48, image.SourceHeight);
        Assert.Equal(2400, image.DesiredWidthTwips);
        Assert.Equal(1200, image.DesiredHeightTwips);

        string rtf = document.ToRtf();
        Assert.Contains(@"\picw96\pich48\picwgoal2400\pichgoal1200", rtf, StringComparison.Ordinal);

        RtfImage roundTripImage = Assert.IsType<RtfImage>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.Equal(96, roundTripImage.SourceWidth);
        Assert.Equal(48, roundTripImage.SourceHeight);
        Assert.Equal(2400, roundTripImage.DesiredWidthTwips);
        Assert.Equal(1200, roundTripImage.DesiredHeightTwips);
    }

    [Fact]
    public void Html_ToRtfDocument_Uses_Shared_Image_Source_Resolution() {
        const string html = "<p><img src=\"https://example.test/chart.png\" data-src=\"data:image/jpeg;base64,/9g=\" alt=\"Chart\"></p>";

        RtfDocument document = html.ToRtfDocument();

        RtfImage image = Assert.Single(Assert.Single(document.Paragraphs).Inlines.OfType<RtfImage>());
        Assert.Equal(RtfImageFormat.Jpeg, image.Format);
        Assert.Equal("Chart", image.Description);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Image_Dimensions() {
        RtfDocument document = RtfDocument.Create();
        RtfImage image = document.AddImage(RtfImageFormat.Png, new byte[] { 0x89, 0x50, 0x4E, 0x47 });
        image.Description = "Chart";
        image.SourceWidth = 96;
        image.SourceHeight = 48;
        image.DesiredWidthTwips = 1440;
        image.DesiredHeightTwips = 720;

        string html = document.ToHtml();

        Assert.Equal("<img src=\"data:image/png;base64,iVBORw==\" alt=\"Chart\" width=\"96\" height=\"48\" style=\"width:72pt;height:36pt;\">", html);

        RtfImage roundTripImage = Assert.Single(html.ToRtfDocument().Paragraphs[0].Inlines.OfType<RtfImage>());
        Assert.Equal(96, roundTripImage.SourceWidth);
        Assert.Equal(48, roundTripImage.SourceHeight);
        Assert.Equal(1440, roundTripImage.DesiredWidthTwips);
        Assert.Equal(720, roundTripImage.DesiredHeightTwips);
    }
}
