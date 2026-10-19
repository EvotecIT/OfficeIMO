using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelXlsbRelatedPartsTests {
    private static string Fixture => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xlsb");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_chart_and_VBA_graphs_survive_worksheet_edits(bool asynchronous) {
        byte[] original = File.ReadAllBytes(Fixture);
        using var document = ExcelDocument.Load(Fixture);
        Assert.Equal(10, document.Sheets.Count);
        Assert.Equal(4, document.WorkbookPartRoot.ChartsheetParts.Count());
        Assert.True(document.HasMacros);
        Assert.Contains(document.ReadVbaProject()!.Modules, item => item.Name == "PreservationNotes");
        document.Sheets[0].CellValue(1, 1, "Edited inert preservation value");
        using var destination = new MemoryStream();
        if (asynchronous) await document.SaveAsync(destination, ExcelFileFormat.Xlsb);
        else document.Save(destination, ExcelFileFormat.Xlsb);
        byte[] saved = destination.ToArray();
        using var before = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
        using var after = new ZipArchive(new MemoryStream(saved), ZipArchiveMode.Read);
        Assert.Equal(before.Entries.Select(item => item.FullName).OrderBy(item => item), after.Entries.Select(item => item.FullName).OrderBy(item => item));
        foreach (var entry in before.Entries.Where(item => item.FullName != "xl/worksheets/sheet1.bin"))
            Assert.Equal(Read(entry), Read(after.GetEntry(entry.FullName)!));
        Assert.NotEqual(Read(before.GetEntry("xl/worksheets/sheet1.bin")!), Read(after.GetEntry("xl/worksheets/sheet1.bin")!));
        using var reloaded = ExcelDocument.Load(new MemoryStream(saved));
        Assert.Equal("Edited inert preservation value", reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(4, reloaded.WorkbookPartRoot.ChartsheetParts.Count());
        Assert.True(reloaded.HasMacros);
        Assert.Equal(saved, document.ToBytes(ExcelFileFormat.Xlsb));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Changed_chart_or_VBA_parts_are_rejected_before_destination_changes(bool macro) {
        using var document = ExcelDocument.Load(Fixture);
        if (macro) {
            using var stream = document.WorkbookPartRoot.VbaProjectPart!.GetStream(FileMode.Create, FileAccess.Write);
            stream.WriteByte(1);
        } else document.WorkbookPartRoot.ChartsheetParts.First().Chartsheet!.GetFirstChild<ChartSheetProperties>()!.CodeName = "ChangedHost";
        byte[] sentinel = { 1, 2, 3, 4 };
        using var destination = new MemoryStream(); destination.Write(sentinel, 0, sentinel.Length);
        Assert.Throws<NotSupportedException>(() => document.Save(destination, ExcelFileFormat.Xlsb));
        Assert.Equal(sentinel, destination.ToArray());
    }

    [Fact]
    public void Direct_worksheet_DOM_changes_do_not_take_the_unchanged_copy_path() {
        using var document = ExcelDocument.Load(Fixture);
        var cell = document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().First(item => item.CellReference?.Value == "A1");
        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
        cell.CellFormula = null; cell.CellValue = null; cell.InlineString = new InlineString(new Text("Direct DOM edit"));
        byte[] saved = document.ToBytes(ExcelFileFormat.Xlsb);
        using var reloaded = ExcelDocument.Load(new MemoryStream(saved));
        Assert.Equal("Direct DOM edit", reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Fact]
    public void Direct_workbook_DOM_changes_are_validated_before_source_copy() {
        using var document = ExcelDocument.Load(Fixture);
        document.WorkbookRoot.GetFirstChild<WorkbookProperties>()!.CodeName = "ChangedWorkbookHost";
        Assert.Throws<NotSupportedException>(() => document.ToBytes(ExcelFileFormat.Xlsb));
    }

    [Fact]
    public void Related_part_graph_obeys_configured_count_and_depth_budgets() {
        byte[] source = File.ReadAllBytes(Fixture);
        Assert.Throws<InvalidDataException>(() => XlsbWorkbookReader.Load(source, new XlsbImportOptions { MaxRelatedParts = 1 }));
        using var extended = new MemoryStream(); extended.Write(source, 0, source.Length);
        using (var package = new ZipArchive(extended, ZipArchiveMode.Update, leaveOpen: true)) {
            using (var writer = new StreamWriter(package.CreateEntry("xl/charts/_rels/chart1.xml.rels").Open()))
                writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"custom\" Type=\"https://example.test/chart-metadata\" Target=\"metadata.xml\"/></Relationships>");
            using var metadata = new StreamWriter(package.CreateEntry("xl/charts/metadata.xml").Open()); metadata.Write("<metadata/>");
        }
        Assert.Throws<InvalidDataException>(() => XlsbWorkbookReader.Load(extended.ToArray(), new XlsbImportOptions { MaxRelationshipDepth = 1 }));
    }

    [Fact]
    public void Direct_shared_string_DOM_edit_is_written_as_a_cell_change() {
        using var document = ExcelDocument.Load(Fixture);
        var cell = document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().First(item => item.DataType?.Value == CellValues.SharedString);
        string reference = cell.CellReference!.Value!;
        int index = int.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture);
        document.WorkbookPartRoot.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ElementAt(index).GetFirstChild<Text>()!.Text = "Edited shared-string content";
        using var reloaded = ExcelDocument.Load(new MemoryStream(document.ToBytes(ExcelFileFormat.Xlsb)));
        Assert.Equal("Edited shared-string content", reloaded.Sheets[0].GetCellText(reloaded.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == reference)));
    }

    [Fact]
    public void Chart_graph_hyperlinks_project_without_activation_and_keep_relationship_ids() {
        using var source = new MemoryStream(); byte[] original = File.ReadAllBytes(Fixture); source.Write(original, 0, original.Length);
        using (var package = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true)) {
            using var writer = new StreamWriter(package.CreateEntry("xl/charts/_rels/chart1.xml.rels").Open());
            writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"chartLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" TargetMode=\"External\" Target=\"https://example.test/chart\"/></Relationships>");
        }
        using var document = ExcelDocument.Load(new MemoryStream(source.ToArray()));
        var link = Assert.Single(document.WorkbookPartRoot.ChartsheetParts.SelectMany(part => part.DrawingsPart!.ChartParts).SelectMany(part => part.HyperlinkRelationships));
        Assert.Equal("chartLink", link.Id); Assert.Equal("https://example.test/chart", link.Uri.OriginalString);
        Assert.Equal(source.ToArray(), document.ToBytes(ExcelFileFormat.Xlsb));
    }

    private static byte[] Read(ZipArchiveEntry entry) {
        using var input = entry.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
    }
}
