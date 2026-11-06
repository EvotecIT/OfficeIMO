using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelSpreadsheetXmlPreservationTests {
    private static readonly XNamespace Ss = "urn:schemas-microsoft-com:office:spreadsheet";
    private static readonly XNamespace X = "urn:schemas-microsoft-com:office:excel";
    private static byte[] Xml(string content) => Encoding.UTF8.GetBytes("<Workbook xmlns='" + Ss + "' xmlns:ss='" + Ss + "' xmlns:x='" + X + "'>" + content + "</Workbook>");

    [Theory]
    [InlineData(105.5, 7, 20.140625, 105.75)]
    [InlineData(7.875, 8, 1.375, 8.25)]
    [InlineData(0, 7, 0, 0)]
    public void Physical_column_width_quantization_is_stable_after_saved_xlsx_and_xml_reimport(
        double points, double digitPixels, double storedWidth, double serializedPoints) {
        string sourcePoints = points.ToString(System.Globalization.CultureInfo.InvariantCulture);
        byte[] source = Xml("<Worksheet ss:Name='Data'><Table ss:DefaultColumnWidth='" + sourcePoints
            + "'><Column ss:Width='" + sourcePoints + "'/></Table></Worksheet>");
        var options = new ExcelSpreadsheetXmlOptions { DefaultFontMaximumDigitWidth = OfficeLength.Pixels(digitPixels) };
        using ExcelDocument document = ExcelDocument.ImportSpreadsheetXml(source, options).Value;
        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Equal(storedWidth, reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single()
            .Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single().Width!.Value);
        Assert.Equal(storedWidth, reopened.OpenXmlDocument.WorkbookPart.WorksheetParts.Single()
            .Worksheet.GetFirstChild<SheetFormatProperties>()!.DefaultColumnWidth!.Value);
        byte[] exported = reopened.ExportSpreadsheetXml(options).Value;
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(exported));
        Assert.Equal(serializedPoints, (double)xml.Descendants(Ss + "Column").Single().Attribute(Ss + "Width")!);
        Assert.Equal(serializedPoints, (double)xml.Descendants(Ss + "Table").Single().Attribute(Ss + "DefaultColumnWidth")!);
        using ExcelDocument reimported = ExcelDocument.ImportSpreadsheetXml(exported, options).Value;
        Assert.Equal(storedWidth, reimported.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single()
            .Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single().Width!.Value);
        Assert.Equal(storedWidth, reimported.OpenXmlDocument.WorkbookPart.WorksheetParts.Single()
            .Worksheet.GetFirstChild<SheetFormatProperties>()!.DefaultColumnWidth!.Value);
    }

    [Fact]
    public void Xml_export_preserves_implicit_zero_width_and_explicit_column_geometry() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        for (int index = 1; index <= 5; index++) sheet.Cell(1, index).SetValue("Column " + index);
        document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet.AddChild(new Columns(
            new Column { Min = 1, Max = 1, Width = 0, CustomWidth = false },
            new Column { Min = 2, Max = 2, Width = 0 },
            new Column { Min = 3, Max = 3, Width = 0, CustomWidth = true },
            new Column { Min = 4, Max = 4, Width = 0, CustomWidth = false, Hidden = true },
            new Column { Min = 5, Max = 5, Width = 8, CustomWidth = false }), true);
        using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        var options = new ExcelSpreadsheetXmlOptions { DefaultFontMaximumDigitWidth = OfficeLength.Pixels(8) };

        byte[] exported = loaded.ExportSpreadsheetXml(options).Value;
        var columns = XDocument.Parse(Encoding.UTF8.GetString(exported)).Descendants(Ss + "Column").ToArray();
        Assert.Equal(new string?[] { null, null, "0", null, "48" },
            columns.Select(column => (string?)column.Attribute(Ss + "Width")));
        Assert.Equal(new string?[] { null, null, null, "1", null },
            columns.Select(column => (string?)column.Attribute(Ss + "Hidden")));

        using var imported = ExcelDocument.ImportSpreadsheetXml(exported, options).Value;
        var importedColumns = imported.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single()
            .Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        Assert.Equal(new double?[] { null, null, 0, null, 8 },
            importedColumns.Select(column => column.Width?.Value));
        Assert.Equal(new[] { false, false, true, false, true },
            importedColumns.Select(column => column.CustomWidth?.Value == true));
        Assert.Equal(new[] { false, false, false, true, false },
            importedColumns.Select(column => column.Hidden?.Value == true));
    }

    [Fact]
    public void Xml_export_preserves_implicit_zero_height_and_explicit_row_geometry() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        for (int index = 1; index <= 5; index++) sheet.Cell(index, 1).SetValue("Row " + index);
        var rows = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single()
            .Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
        for (int index = 0; index < rows.Length; index++) {
            rows[index].Height = index == 4 ? 8 : 0;
            rows[index].CustomHeight = index == 2;
        }
        rows[1].CustomHeight = null;
        rows[3].Hidden = true;
        using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));

        byte[] exported = loaded.ExportSpreadsheetXml().Value;
        var xmlRows = XDocument.Parse(Encoding.UTF8.GetString(exported)).Descendants(Ss + "Row").ToArray();
        Assert.Equal(new string?[] { null, null, "0", null, "8" },
            xmlRows.Select(row => (string?)row.Attribute(Ss + "Height")));
        Assert.Equal(new string?[] { null, null, null, "1", null },
            xmlRows.Select(row => (string?)row.Attribute(Ss + "Hidden")));

        using var imported = ExcelDocument.ImportSpreadsheetXml(exported).Value;
        var importedRows = imported.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single()
            .Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
        Assert.Equal(new double?[] { null, null, 0, null, 8 },
            importedRows.Select(row => row.Height?.Value));
        Assert.Equal(new[] { false, false, true, false, true },
            importedRows.Select(row => row.CustomHeight?.Value == true));
        Assert.Equal(new[] { false, false, false, true, false },
            importedRows.Select(row => row.Hidden?.Value == true));
    }

    [Fact]
    public void Automatic_font_metrics_report_the_physical_column_width_approximation() {
        byte[] source = Xml("<Styles><Style ss:ID='Default'><Font ss:FontName='Calibri' ss:Size='11'/></Style></Styles>"
            + "<Worksheet ss:Name='Data'><Table><Column ss:Width='105.5'/></Table></Worksheet>");
        var imported = ExcelDocument.ImportSpreadsheetXml(source);
        using ExcelDocument document = imported.Value;
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "excel-xml-approximation"
            && item.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Contains(document.ExportSpreadsheetXml().Report.Diagnostics, item => item.Code == "excel-xml-approximation"
            && item.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Native_Excel_geometry_comments_and_font_are_imported_and_reexported() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "SpreadsheetXmlCorpus", "excel-geometry.xml");
        using var document = ExcelDocument.ImportSpreadsheetXml(path, new() { DefaultFontMaximumDigitWidth = OfficeIMO.OfficeLength.Points(6) }).Value;
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
        Assert.Equal("C10:D11", part.Worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Single().Reference!.Value);
        Assert.Equal("Native XML note", part.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>().Single().CommentText!.InnerText);
        Assert.Equal(new[] { 8d, 16d, 32d }, part.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Where(item => item.Min!.Value is >= 3 and <= 5).Select(item => item.Width!.Value));
        var xml = XDocument.Parse(Encoding.UTF8.GetString(document.ExportSpreadsheetXml(new() { DefaultFontMaximumDigitWidth = OfficeIMO.OfficeLength.Pixels(8) }).Value));
        Assert.Equal(new[] { "48", "96", "192" }, xml.Descendants(Ss + "Column").Where(item => (int?)item.Attribute(Ss + "Index") is >= 3 and <= 5).Select(item => (string?)item.Attribute(Ss + "Width")));
        Assert.Empty(document.ValidateOpenXml());
    }

    [Theory]
    [InlineData(-1L)]
    [InlineData(0L)]
    [InlineData(9524L)]
    [InlineData(39014401L)]
    public void MeasuredFontWidthRejectsOutOfRangePhysicalValuesBeforeReading(long emus) {
        byte[] source = Xml("<Worksheet ss:Name='Data'><Table/></Worksheet>");
        var options = new ExcelSpreadsheetXmlOptions {
            DefaultFontMaximumDigitWidth = OfficeIMO.OfficeLength.Emus(emus)
        };
        var error = Assert.Throws<ArgumentOutOfRangeException>(() =>
            ExcelDocument.ImportSpreadsheetXml(source, options));
        Assert.Equal(nameof(ExcelSpreadsheetXmlOptions.DefaultFontMaximumDigitWidth), error.ParamName);
    }

    [Fact]
    public void Inherited_appearance_dates_names_links_merges_and_print_settings_survive_xml_exchange() {
        byte[] source = Xml("""
            <Styles>
             <Style ss:ID="Default"><Font ss:FontName="Calibri" ss:Size="11"/><Alignment ss:Vertical="Bottom"/></Style>
             <Style ss:ID="base"><Font ss:Bold="1" ss:Color="#123456"/><Interior ss:Pattern="Solid" ss:Color="#FEDCBA"/>
              <Borders><Border ss:Position="Bottom" ss:LineStyle="Dash" ss:Weight="2" ss:Color="#010203"/></Borders></Style>
             <Style ss:ID="child" ss:Parent="base"><Font ss:Italic="1"/><Alignment ss:WrapText="1" ss:Rotate="-45"/><Protection ss:Protected="0" x:HideFormula="1"/></Style>
            </Styles>
            <Names><NamedRange ss:Name="GlobalValue" ss:RefersTo="=Data!R1C1"/></Names>
            <Worksheet ss:Name="Data"><Names><NamedRange ss:Name="Print_Area" ss:RefersTo="=Data!R1C1:R3C3"/></Names>
             <Table ss:DefaultRowHeight="18"><Column ss:Index="2" ss:StyleID="base" ss:Hidden="1"/>
              <Row><Cell ss:StyleID="child" ss:HRef="https://example.org/a?x=1&amp;y=2" x:HRefScreenTip="Open">
               <Data ss:Type="DateTime">2026-10-19T12:30:00</Data><Comment ss:Author="A"><Data>Note</Data></Comment></Cell><Cell><Data ss:Type="String">Column style</Data></Cell></Row>
              <Row ss:Index="3"><Cell ss:MergeAcross="1" ss:MergeDown="1"><Data ss:Type="String">Merged</Data></Cell></Row>
             </Table><x:WorksheetOptions><x:PageSetup><x:Layout x:Orientation="Landscape" x:CenterHorizontal="1"/><x:PageMargins x:Left="0.4" x:Right="0.5" x:Top="0.6" x:Bottom="0.7"/><x:Header x:Margin="0.2" x:Data="&amp;CReport"/></x:PageSetup><x:Print><x:PaperSizeIndex>9</x:PaperSizeIndex><x:FitWidth>1</x:FitWidth><x:FitHeight>0</x:FitHeight></x:Print><x:FitToPage/><x:DoNotDisplayGridlines/><x:Zoom>125</x:Zoom></x:WorksheetOptions>
            </Worksheet>
            """);
        var documentResult = ExcelDocument.ImportSpreadsheetXml(source);
        documentResult.ThrowIfLoss();
        using var document = documentResult.Value;
        using var reopened = ExcelDocument.ImportSpreadsheetXml(document.ExportSpreadsheetXml().Value).Value;
        foreach (var workbook in new[] { document, reopened }) {
            var book = workbook.OpenXmlDocument.WorkbookPart!; var sheet = book.WorksheetParts.Single().Worksheet!;
            var style = book.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)sheet.Descendants<Cell>().Single(item => item.CellReference == "A1").StyleIndex!.Value);
            var font = book.WorkbookStylesPart.Stylesheet.Fonts!.Elements<Font>().ElementAt((int)style.FontId!.Value);
            Assert.True(font.Bold!.Val!.Value); Assert.True(font.Italic!.Val!.Value); Assert.Equal("FF123456", font.Color!.Rgb!.Value);
            Assert.NotEqual(0U, style.NumberFormatId!.Value); Assert.True(style.Alignment!.WrapText!.Value); Assert.Equal(135U, style.Alignment.TextRotation!.Value);
            var fill = book.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)style.FillId!.Value);
            Assert.Equal(PatternValues.Solid, fill.PatternFill!.PatternType!.Value);
            Assert.Equal("FFFEDCBA", fill.PatternFill.ForegroundColor!.Rgb!.Value);
            Assert.False(style.Protection!.Locked!.Value); Assert.True(style.Protection.Hidden!.Value);
            Assert.Equal("A3:B4", sheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Single().Reference!.Value);
            Assert.Equal(2, book.Workbook.DefinedNames!.ChildElements.Count);
            Assert.Equal("Data!$A$1", book.Workbook.DefinedNames.Elements<DefinedName>().Single(item => item.Name == "GlobalValue").Text);
            Assert.Equal("https://example.org/a?x=1&y=2", book.WorksheetParts.Single().HyperlinkRelationships.Single().Uri.OriginalString);
            Assert.Equal(OrientationValues.Landscape, sheet.GetFirstChild<PageSetup>()!.Orientation!.Value);
            Assert.Equal(.4, sheet.GetFirstChild<PageMargins>()!.Left!.Value); Assert.Equal("&CReport", sheet.GetFirstChild<HeaderFooter>()!.OddHeader!.Text);
            Assert.Empty(workbook.ValidateOpenXml());
        }
    }

    [Theory]
    [InlineData("excel-preservation.xml")]
    [InlineData("officeimo-preservation.xml")]
    public void Producer_and_consumer_controls_retain_supported_appearance_and_metadata(string fixture) {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "SpreadsheetXmlCorpus", fixture);
        using var document = ExcelDocument.ImportSpreadsheetXml(path, new() { DefaultFontMaximumDigitWidth = OfficeIMO.OfficeLength.Pixels(8) }).Value;
        var book = document.OpenXmlDocument.WorkbookPart!; var part = book.WorksheetParts.Single();
        var cell = part.Worksheet.Descendants<Cell>().Single(item => item.CellReference == "A1");
        var style = book.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        var font = book.WorkbookStylesPart.Stylesheet.Fonts!.Elements<Font>().ElementAt((int)style.FontId!.Value);
        Assert.True(font.Bold!.Val!.Value); Assert.True(font.Italic!.Val!.Value); Assert.Equal("FF123456", font.Color!.Rgb!.Value);
        Assert.Equal(PatternValues.Solid, book.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)style.FillId!.Value).PatternFill!.PatternType!.Value);
        Assert.True(style.Protection!.Hidden!.Value); Assert.Equal("Open", part.Worksheet.GetFirstChild<Hyperlinks>()!.Elements<Hyperlink>().Single().Tooltip!.Value);
        Assert.Equal(2, book.Workbook.DefinedNames!.ChildElements.Count); Assert.Empty(document.ValidateOpenXml());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unstyled_cells_keep_effective_column_or_row_styles(bool rowStyle) {
        using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Inherited"); sheet.Cell(1, 1).SetBold();
        var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var cell = worksheet.Descendants<Cell>().Single(); uint id = cell.StyleIndex!.Value; cell.StyleIndex = null;
        if (rowStyle) { var row = worksheet.Descendants<Row>().Single(); row.StyleIndex = id; row.CustomFormat = true; }
        else worksheet.AddChild(new Columns(new Column { Min = 1, Max = 1, Style = id }), true);
        using var copy = ExcelDocument.ImportSpreadsheetXml(document.ExportSpreadsheetXml().Value).Value;
        var book = copy.OpenXmlDocument.WorkbookPart!; uint style = book.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single().StyleIndex!.Value;
        var format = book.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)style);
        Assert.True(book.WorkbookStylesPart.Stylesheet.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value).Bold!.Val!.Value);
    }

    [Fact]
    public void Unmapped_nested_metadata_is_visible_in_export_diagnostics() {
        using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Data"); sheet.Cell(1, 1).SetValue("Protected");
        sheet.SetTabColor(global::OfficeIMO.OfficeColor.Parse("FF0000")); sheet.Protect(new() { AllowSelectLockedCells = false });
        var result = document.ExportSpreadsheetXml();
        Assert.Contains(result.Report.Diagnostics, item => item.Message.Contains("tabColor"));
        using var copy = ExcelDocument.ImportSpreadsheetXml(result.Value).Value;
        Assert.True(copy.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProtection>()!.SelectLockedCells!.Value);
    }

    [Fact]
    public void Styleless_fast_package_export_obeys_cell_budget_and_preserves_values() {
        using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("first"); sheet.Cell(2, 1).SetValue("second");
        using var copy = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Contains("MaxCells", Assert.Throws<InvalidDataException>(() => copy.ExportSpreadsheetXml(new() { MaxCells = 1 })).Message);
        using var xml = ExcelDocument.ImportSpreadsheetXml(copy.ExportSpreadsheetXml().Value).Value;
        Assert.Equal("first", xml.Sheets[0].Cell(1, 1).GetValue<string>());
    }

    [Fact]
    public void Rich_text_cell_and_comment_runs_keep_the_supported_html_emphasis_and_fonts() {
        byte[] xml = Xml("<Worksheet ss:Name='Data'><Table><Row><Cell><Data ss:Type='String' xmlns:h='http://www.w3.org/TR/REC-html40'>Plain <h:B>bold <h:I>italic</h:I></h:B><h:Font h:Face='Arial' h:Size='14' h:Color='#FF0000'>red</h:Font></Data><Comment ss:Author='A'><Data xmlns:h='http://www.w3.org/TR/REC-html40'><h:U>Note</h:U></Data></Comment></Cell></Row></Table></Worksheet>");
        var documentResult = ExcelDocument.ImportSpreadsheetXml(xml);
        documentResult.ThrowIfLoss();
        using var document = documentResult.Value;
        using var copy = ExcelDocument.ImportSpreadsheetXml(document.ExportSpreadsheetXml().Value).Value;
        foreach (var current in new[] { document, copy }) {
            var part = current.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
            Run[] runs = part.Worksheet.Descendants<Cell>().Single().InlineString!.Elements<Run>().ToArray();
            Assert.Equal("Plain bold italicred", string.Concat(runs.Select(item => item.InnerText)));
            Assert.Contains(runs, item => item.RunProperties?.GetFirstChild<Bold>()?.Val?.Value == true && item.RunProperties.GetFirstChild<Italic>()?.Val?.Value == true);
            Assert.Contains(runs, item => item.RunProperties?.GetFirstChild<RunFont>()?.Val?.Value == "Arial" && item.RunProperties.GetFirstChild<FontSize>()?.Val?.Value == 14);
            Assert.NotNull(part.WorksheetCommentsPart!.Comments!.Descendants<RunProperties>().Single().GetFirstChild<Underline>());
            Assert.Empty(current.ValidateOpenXml());
        }
    }

    [Fact]
    public void Native_Excel_readback_retains_rich_cell_runs_and_underlined_comment() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "SpreadsheetXmlCorpus", "excel-rich-text-readback.xlsx");
        using var document = ExcelDocument.Load(path);
        var part = document.OpenXmlDocument.WorkbookPart!;
        var runs = part.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().Single().Elements<Run>().ToArray();
        Assert.Equal("Plain bold italic red", string.Concat(runs.Select(item => item.InnerText)));
        Assert.NotNull(runs[1].RunProperties!.GetFirstChild<Bold>());
        Assert.NotNull(runs[2].RunProperties!.GetFirstChild<Italic>());
        Assert.Equal("Arial", runs[3].RunProperties!.GetFirstChild<RunFont>()!.Val!.Value);
        Assert.Equal(14d, runs[3].RunProperties!.GetFirstChild<FontSize>()!.Val!.Value);
        Assert.Equal(10U, runs[3].RunProperties!.GetFirstChild<Color>()!.Indexed!.Value);
        Assert.NotNull(part.WorksheetParts.Single().WorksheetCommentsPart!.Comments!.Descendants<RunProperties>().Single().GetFirstChild<Underline>());
    }

    [Theory]
    [InlineData("<Row><Cell ss:MergeAcross='1'/><Cell ss:Index='2'/></Row>")]
    [InlineData("<Row><Cell ss:MergeDown='1'/></Row><Row><Cell/></Row>")]
    [InlineData("<Row><Cell ss:MergeDown='65536'/></Row>")]
    public void Overlapping_or_out_of_bounds_xml_merges_fail_without_silently_losing_cells(string rows) =>
        Assert.Throws<InvalidDataException>(() => ExcelDocument.ImportSpreadsheetXml(Xml("<Worksheet ss:Name='Data'><Table>" + rows + "</Table></Worksheet>")));

    [Fact]
    public void Export_keeps_empty_merge_anchors_and_rejects_covered_cell_data() {
        using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Data");
        var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        worksheet.AddChild(new MergeCells(new MergeCell { Reference = "C3:D4" }), true);
        using var copy = ExcelDocument.ImportSpreadsheetXml(document.ExportSpreadsheetXml().Value).Value;
        Assert.Equal("C3:D4", copy.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Single().Reference!.Value);
        sheet.Cell(4, 4).SetValue("Cannot discard");
        Assert.Throws<InvalidDataException>(() => document.ExportSpreadsheetXml());
    }
}
