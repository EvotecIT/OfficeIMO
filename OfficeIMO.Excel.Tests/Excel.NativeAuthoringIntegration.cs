using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

[CollectionDefinition("Native workbook authoring", DisableParallelization = true)]
public sealed class NativeWorkbookAuthoringCollection { }

// Existing AutoFit failure checks observe the process-wide native spool directory.
[Collection("Native workbook authoring")]
public sealed class ExcelNativeAuthoringIntegrationTests {
    [Fact]
    public void ApplyingThemeRetainsNativeLinksAndChangesSubsequentStatusDefaults() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Theme");
        sheet.Cell(1, 1).SetValue("Linked").SetColor(OfficeColor.Accent1);
        var theme = OfficeTheme.Dark.WithPalette(OfficeTheme.Dark.Palette.With(
            accent1: OfficeColor.Parse("#123456"), danger: OfficeColor.Parse("#654321")));
        document.ApplyTheme(theme);
        sheet.Cell(2, 1).SetValue("Status").SetFillColor(OfficeColor.Danger);
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        var styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
        var part = package.WorkbookPart.WorksheetParts.Single();
        var first = Format(styles, Cell(part, "A1"));
        Assert.Equal(4U, styles.Fonts!.Elements<Font>().ElementAt((int)first.FontId!.Value).Color!.Theme!.Value);
        var second = Format(styles, Cell(part, "A2"));
        Assert.Equal("FF654321", styles.Fills!.Elements<Fill>().ElementAt((int)second.FillId!.Value).PatternFill!.ForegroundColor!.Rgb!.Value);
        Assert.Contains("123456", package.WorkbookPart.ThemePart!.Theme.OuterXml);
        Assert.Equal(theme.Name, package.WorkbookPart.ThemePart.Theme.Name!.Value);
    }

    [Fact]
    public void CrossSheetPivotCachesSourceRowsAndPlacesThePivotOnItsTarget() {
        using var document = ExcelDocument.Create();
        var source = document.AddSheet("Input");
        source.Cell(1, 1).SetValue("Group"); source.Cell(1, 2).SetValue("Amount");
        source.Cell(2, 1).SetValue("North"); source.Cell(2, 2).SetValue(10);
        source.Cell(3, 1).SetValue("South"); source.Cell(3, 2).SetValue(20);
        var target = document.AddSheet("Pivot");
        target.Cell(1, 1).SetValue("Wrong target headers"); target.Cell(2, 2).SetValue(999);
        target.AddPivotTable(source, "A1:B3", "D4", new ExcelPivotTableOptions {
            Name = "Summary", RowFields = new[] { "Group" },
            DataFields = new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum) }, SaveSourceData = true
        });
        Assert.Equal("Input", target.GetPivotTables().Single().SourceSheet);
        Assert.Empty(source.GetPivotTables());
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        var pivot = package.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
        Assert.Equal("Input", pivot.PivotTableCacheDefinitionPart!.PivotCacheDefinition.CacheSource!.WorksheetSource!.Sheet!.Value);
        Assert.Equal("A1:B3", pivot.PivotTableCacheDefinitionPart.PivotCacheDefinition.CacheSource.WorksheetSource.Reference!.Value);
        var records = pivot.PivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Single().PivotCacheRecords;
        Assert.Equal(new[] { 10D, 20D }, records.Descendants<NumberItem>().Select(item => item.Val!.Value));
        using var foreign = ExcelDocument.Create();
        Assert.Throws<ArgumentException>(() => target.AddPivotTable(foreign.AddSheet("Input"), "A1:B3", "H4", null));
    }

    [Fact]
    public void ClearingNamedTableStylePreservesItsFiltersAndVisualFlags() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Table");
        sheet.Cell(1, 1).SetValue("Id"); sheet.Cell(2, 1).SetValue(7);
        sheet.AddTable("A1:A2", name: "Data");
        sheet.SetTableStyle("Data", ExcelTableStyle.None, showFirstColumn: true, showRowStripes: false);
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        var table = package.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table;
        Assert.Null(table.TableStyleInfo!.Name);
        Assert.True(table.TableStyleInfo.ShowFirstColumn!.Value);
        Assert.False(table.TableStyleInfo.ShowRowStripes!.Value);
        Assert.NotNull(table.AutoFilter);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void IncrementalRowsPreserveTheFinalSchemaTypesLayoutAndMetadata(bool headers) {
        using var output = new MemoryStream();
        var options = new ExcelWriterOptions {
            Address = "C3", Title = "Orders", BoldHeaders = true, FreezeHeader = true, FreezeFirstColumn = true,
            CreateTable = true, IncludeHeaders = headers, TableStyle = ExcelTableStyle.None, AutoFit = true,
            ColumnFormat = new ExcelColumnFormatPlan().AddFormat("Amount", "0.000"),
            DocumentProperties = new Dictionary<string, object?> { ["Author"] = "Owner", ["Title"] = "Native export", ["Company"] = "Example", ["Created"] = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc) },
            SaveOptions = new ExcelSaveOptions { ValidateOpenXml = true }
        };
        ExcelDataSetImportResult result;
        using (var workbook = ExcelDocument.CreateWriter(output, options, leaveOpen: true)) {
            using var rows = workbook.BeginTable("Orders");
            object?[] first = { "First", true };
            rows.WriteRow(new[] { "Item", "Active" }, first);
            first[0] = "Mutated after ingestion";
            rows.WriteRow(new[] { "Item", "Amount", "When" }, new object?[] { "Second", 42.125m, new DateTime(2025, 1, 2, 12, 0, 0) });
            result = rows.Complete();
        }
        int dataRow = headers ? 5 : 4;
        Assert.Equal(headers ? "C4:F6" : "C4:F5", result.Range);
        Assert.Equal((2, 4), (result.RowCount, result.ColumnCount));
        using var package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var workbookPart = package.WorkbookPart!;
        var part = workbookPart.WorksheetParts.Single();
        var styles = workbookPart.WorkbookStylesPart!.Stylesheet;
        Assert.Equal("Orders", Cell(part, "C3").InnerText);
        Assert.True(IsBold(styles.Fonts!.Elements<Font>().ElementAt((int)Format(styles, Cell(part, "C3")).FontId!.Value)));
        Assert.Equal("First", Cell(part, "C" + dataRow).InnerText);
        Assert.Equal(CellValues.Boolean, Cell(part, "D" + dataRow).DataType!.Value);
        Assert.Null(Cell(part, "E" + dataRow).CellValue);
        Assert.Null(Cell(part, "D" + (dataRow + 1)).CellValue);
        Assert.Equal("42.125", Cell(part, "E" + (dataRow + 1)).CellValue!.Text);
        Assert.Equal(new DateTime(2025, 1, 2, 12, 0, 0).ToOADate().ToString(CultureInfo.InvariantCulture), Cell(part, "F" + (dataRow + 1)).CellValue!.Text);
        Assert.Equal("0.000", styles.NumberingFormats!.Elements<NumberingFormat>().Single(format => format.NumberFormatId!.Value == Format(styles, Cell(part, "E" + (dataRow + 1))).NumberFormatId!.Value).FormatCode!.Value);
        var table = part.TableDefinitionParts.Single().Table;
        Assert.Equal(headers ? 1U : 0U, table.HeaderRowCount!.Value);
        Assert.Equal(result.Range, table.Reference!.Value);
        Assert.Null(table.TableStyleInfo!.Name);
        Assert.Equal(headers, table.AutoFilter != null);
        var pane = part.Worksheet.Descendants<Pane>().Single();
        Assert.Equal(3D, pane.HorizontalSplit!.Value);
        Assert.Equal(headers ? 4D : 0D, pane.VerticalSplit!.Value);
        Assert.Equal("Owner", package.PackageProperties.Creator);
        Assert.Equal("Native export", package.PackageProperties.Title);
        Assert.Equal("Example", package.ExtendedFilePropertiesPart!.Properties.Company!.Text);
        Assert.Equal(new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc), package.PackageProperties.Created!.Value.ToUniversalTime());
        var columns = part.Worksheet.Descendants<Column>().ToArray();
        Assert.Equal(new[] { 3U, 4U, 5U, 6U }, columns.Select(column => column.Min!.Value));
        Assert.All(columns, column => Assert.True(column.Width!.Value > 0));
        if (headers) {
            Assert.Equal(new[] { "Item", "Active", "Amount", "When" }, new[] { "C4", "D4", "E4", "F4" }.Select(reference => Cell(part, reference).InnerText));
            Assert.True(IsBold(styles.Fonts.Elements<Font>().ElementAt((int)Format(styles, Cell(part, "C4")).FontId!.Value)));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IncompleteOrCancelledIncrementalTablesPreserveTheDestination(bool cancel) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Incremental-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "existing.xlsx");
        File.WriteAllText(path, "sentinel");
        try {
            Assert.Throws<IOException>(() => ExcelDocument.CreateWriter(path, new ExcelWriterOptions { NoClobber = true }));
            using var cancellation = new CancellationTokenSource();
            using (var workbook = ExcelDocument.CreateWriter(path)) {
                using var rows = workbook.BeginTable("Data", cancellation.Token);
                rows.WriteRow(new[] { "Id" }, new object?[] { 7L });
                if (cancel) {
                    cancellation.Cancel();
                    Assert.ThrowsAny<OperationCanceledException>(() => rows.Complete());
                }
            }
            Assert.Equal("sentinel", File.ReadAllText(path));
            Assert.Single(Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void DeclaredColumnsKeepAnEmptyTypedSchemaWithoutAddingADataRow() {
        using var output = new MemoryStream();
        using (var workbook = ExcelDocument.CreateWriter(output, leaveOpen: true)) {
            using var rows = workbook.BeginTable("Empty");
            rows.DeclareColumns(new[] { "Id", "Name" });
            var result = rows.Complete();
            Assert.Equal((0, 2), (result.RowCount, result.ColumnCount));
        }
        using var package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        var data = package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetData>()!;
        var header = Assert.Single(data.Elements<Row>());
        Assert.Equal(new[] { "Id", "Name" }, header.Elements<Cell>().Select(cell => cell.InnerText));
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void WorkbookAbortDiscardsEarlierCompletedTables() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Abort-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "existing.xlsx");
        File.WriteAllText(path, "sentinel");
        try {
            using (var workbook = ExcelDocument.CreateWriter(path)) {
                using var first = workbook.BeginTable("First");
                first.WriteRow(new[] { "Id" }, new object?[] { 1 });
                first.Complete();
                workbook.Abort();
            }
            Assert.Equal("sentinel", File.ReadAllText(path));
            Assert.Single(Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void NoClobberAlsoProtectsADestinationCreatedDuringExport() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-NoClobber-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "new.xlsx");
        try {
            var workbook = ExcelDocument.CreateWriter(path, new ExcelWriterOptions { NoClobber = true });
            using (var rows = workbook.BeginTable("Data")) {
                rows.WriteRow(new[] { "Id" }, new object?[] { 7 });
                rows.Complete();
            }
            File.WriteAllText(path, "created by another writer");
            Assert.Throws<IOException>(() => workbook.Dispose());
            Assert.Equal("created by another writer", File.ReadAllText(path));
            Assert.Single(Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void FinalMetadataMustFitTheTemporaryPackageBudgetBeforePublication() {
        using var baseline = new MemoryStream();
        using (var workbook = ExcelDocument.CreateWriter(baseline, leaveOpen: true)) {
            using var rows = workbook.BeginTable("Data");
            rows.WriteRow(new[] { "Id" }, new object?[] { 7 });
            rows.Complete();
        }
        var random = new Random(42);
        var metadata = new byte[8192];
        random.NextBytes(metadata);
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Budget-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "existing.xlsx");
        File.WriteAllText(path, "sentinel");
        try {
            var workbook = ExcelDocument.CreateWriter(path, new ExcelWriterOptions {
                DocumentProperties = new Dictionary<string, object?> { ["Company"] = Convert.ToBase64String(metadata) },
                SaveOptions = new ExcelSaveOptions { MaxTemporaryPackageBytes = baseline.Length + 256 }
            });
            using (var rows = workbook.BeginTable("Data")) {
                rows.WriteRow(new[] { "Id" }, new object?[] { 7 });
                rows.Complete();
            }
            Assert.Contains("MaxTemporaryPackageBytes", Assert.Throws<InvalidOperationException>(() => workbook.Dispose()).Message);
            Assert.Equal("sentinel", File.ReadAllText(path));
            Assert.Single(Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void ExplicitPasswordPolicyUsesTheNativeSaveOwnerAndItsCapturedOptions() {
        using var output = new MemoryStream();
        var save = new ExcelSaveOptions { Password = "synthetic-writer-password", ValidateOpenXml = true };
        using (var workbook = ExcelDocument.CreateWriter(output, new ExcelWriterOptions { SaveOptions = save }, leaveOpen: true)) {
            save.Password = "changed after writer creation";
            using var rows = workbook.BeginTable("Protected");
            rows.WriteRow(new[] { "Text" }, new object?[] { "retained value" });
            rows.Complete();
        }
        using var restored = ExcelDocument.Load(new MemoryStream(output.ToArray()), new ExcelLoadOptions { Password = "synthetic-writer-password" });
        Assert.Equal("retained value", restored.Sheets.Single().Cell(2, 1).GetValue<string>());
    }

    private static bool IsBold(Font font) => font.Bold?.Val?.Value ?? font.Bold != null;
    private static Cell Cell(WorksheetPart part, string reference) => part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == reference);
    private static CellFormat Format(Stylesheet styles, Cell cell) => styles.CellFormats!.Elements<CellFormat>().ElementAt((int)(cell.StyleIndex?.Value ?? 0U));
}
