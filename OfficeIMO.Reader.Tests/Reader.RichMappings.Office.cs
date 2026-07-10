using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOfficeRichMappingTests {
    [Fact]
    public void DocumentReader_WordRichMapping_UsesInspectionBlocksTablesLinksAndProperties() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.BuiltinDocumentProperties.Title = "Rich policy";
            document.BuiltinDocumentProperties.Creator = "OfficeIMO";
            document.AddParagraph("Policy").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Read ").AddHyperLink("the policy", new Uri("https://example.test/policy"));
            WordTable table = document.AddTable(2, 2);
            table.Title = "Inventory";
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Qty";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Bandage";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "4";
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = DocumentReader.ReadDocument(stream, "policy.docx");

        Assert.Equal("Rich policy", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        Assert.Contains(result.Blocks, block => block.Kind == "heading" && block.Level == 1 && block.Location.HeadingPath == "Policy");
        Assert.Contains(result.Blocks, block => block.Kind == "paragraph" && block.Location.HeadingPath == "Policy");
        ReaderTable mapped = Assert.Single(result.Tables);
        Assert.Equal("Inventory", mapped.Title);
        Assert.Equal("Bandage", mapped.Rows[0][0]);
        Assert.Equal("https://example.test/policy", Assert.Single(result.Links).Uri);
        Assert.Contains("officeimo.word.inspection-snapshot", result.CapabilitiesUsed);
    }

    [Fact]
    public void DocumentReader_ExcelRichMapping_UsesFormalTablesCellLinksAndProperties() {
        using var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create(stream)) {
            document.BuiltinDocumentProperties.Title = "Rich workbook";
            document.BuiltinDocumentProperties.Creator = "OfficeIMO";
            ExcelSheet sheet = document.AddWorkSheet("Inventory");
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Qty");
            sheet.Cell(2, 1, "Bandage");
            sheet.Cell(2, 2, 4);
            sheet.AddTable("A1:B2", hasHeader: true, name: "InventoryTable", style: TableStyle.TableStyleMedium2);
            sheet.SetHyperlink(2, 1, "https://example.test/bandage", display: "Bandage");
            ExcelSheet rawSheet = document.AddWorkSheet("Raw");
            rawSheet.Cell(1, 1, "Metric");
            rawSheet.Cell(1, 2, "Value");
            rawSheet.Cell(2, 1, "Retries");
            rawSheet.Cell(2, 2, 3);
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = DocumentReader.ReadDocument(stream, "inventory.xlsx");

        Assert.Equal("Rich workbook", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        ReaderTable mapped = Assert.Single(result.Tables, table => table.Kind == "excel-table");
        Assert.Equal("InventoryTable", mapped.Title);
        Assert.Equal("excel-table", mapped.Kind);
        Assert.Equal("Bandage", mapped.Rows[0][0]);
        Assert.Contains(result.Tables, table => table.Location?.Sheet == "Raw");
        OfficeDocumentLink link = Assert.Single(result.Links);
        Assert.Equal("Inventory", link.Location.Sheet);
        Assert.Equal("A2", link.Location.A1Range);
        Assert.Equal("https://example.test/bandage", link.Uri);
        OfficeDocumentPage inventoryPage = Assert.Single(result.Pages, page => page.Name == "Inventory");
        Assert.Same(mapped, Assert.Single(inventoryPage.Tables));
        Assert.Contains("officeimo.excel.inspection-snapshot", result.CapabilitiesUsed);
    }

    [Fact]
    public void DocumentReader_PowerPointRichMapping_UsesShapesTablesChartsLinksAndSlideGeometry() {
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            presentation.BuiltinDocumentProperties.Title = "Rich deck";
            presentation.BuiltinDocumentProperties.Creator = "OfficeIMO";
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox title = slide.AddTextBox("Overview");
            title.AddParagraph("Summary");
            title.Paragraphs[0].Runs[0].SetHyperlink("https://example.test/deck");
            PowerPointTable table = slide.AddTable(2, 2);
            table.GetCell(0, 0).Text = "Name";
            table.GetCell(0, 1).Text = "Qty";
            table.GetCell(1, 0).Text = "Bandage";
            table.GetCell(1, 1).Text = "4";
            slide.AddChart(new PowerPointChartData(
                new[] { "Q1", "Q2" },
                new[] { new PowerPointChartSeries("Sales", new[] { 1D, 2D }) }));
            presentation.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = DocumentReader.ReadDocument(stream, "deck.pptx");

        Assert.Equal("Rich deck", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.True(page.Width > 0);
        Assert.True(page.Height > 0);
        Assert.Contains(result.Blocks, block => block.Kind == "paragraph" && block.Text == "Overview" && block.Region != null);
        Assert.Equal("Bandage", Assert.Single(result.Tables).Rows[0][0]);
        Assert.Equal("https://example.test/deck", Assert.Single(result.Links).Uri);
        ReaderVisual chart = Assert.Single(result.Visuals);
        Assert.Equal("chart", chart.Kind);
        Assert.Contains("Sales", chart.Content, StringComparison.Ordinal);
        Assert.Equal("3", Assert.Single(result.Metadata, item => item.Name == "ShapeCount").Value);
        Assert.Contains("officeimo.powerpoint.chart-snapshot", result.CapabilitiesUsed);
    }
}
