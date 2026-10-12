using System.Globalization;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNumbersWritingTests {
    [Fact]
    public void Creation_matches_the_independently_decoded_package_opened_and_edited_by_Numbers() {
        byte[] bytes = Example().SaveBytes();
        Assert.Equal(File.ReadAllBytes(Fixture("numbers-created-v14.5.numbers")), bytes);
        Assert.Equal(bytes, Example().SaveBytes());
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(bytes).ReadNumbers();
        VerifyTables(projection, 12.5);
        IWorkTable table = projection.Sheets[0].Tables[0];
        Assert.Equal(40, table.Geometry!.LeftPoints); Assert.Equal(80, table.Geometry.TopPoints);
        Assert.Equal(520, table.Geometry.WidthPoints); Assert.Equal(144, table.Geometry.HeightPoints);
        Assert.Equal(36, table.RowHeights[4]); Assert.Equal(130, table.ColumnWidths[4]);
        Assert.Equal("Arial", table.GetParagraphStyle(4, 4)!.TextStyle.FontName);
        Assert.Equal(14, table.GetParagraphStyle(4, 4)!.TextStyle.FontSizePoints);
        Assert.Equal("003366", table.GetParagraphStyle(4, 4)!.TextStyle.Color!.RgbHex);
        Assert.Equal("FFFFFF", table.GetFill(4, 4)!.Color!.RgbHex);
        Assert.Equal(4, table.GetPadding(4, 4)!.BottomPoints);
        Assert.Equal(IWorkCellVerticalAlignment.Top, table.GetVerticalAlignment(4, 4));
        Assert.False(table.AutoResizeRows);
        Assert.Equal(new[] { "Data", "Totals" }, projection.Sheets[0].Drawables.Select(drawable => drawable.Table!.Name));
    }

    [Fact]
    public void Native_edited_source_retains_all_tables_and_matches_Apple_saved_Excel_typed_values() {
        string path = Fixture("numbers-created-edited-v14.5.numbers");
        VerifyTables(IWorkSourceDocument.Open(path).ReadNumbers(), 20.75);
        Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly }));
        using var conversion = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED");
        Assert.Throws<InvalidOperationException>(() => conversion.Report.RequireCompleteEditableReconstruction());
        using var output = new MemoryStream(); conversion.Value.Save(output); output.Position = 0;
        using (var saved = SpreadsheetDocument.Open(output, false)) {
            Assert.Empty(new OpenXmlValidator().Validate(saved));
            using var apple = SpreadsheetDocument.Open(Fixture("numbers-created-v14.5.xlsx"), false);
            for (int index = 0; index < 2; index++)
                Assert.Equal(Values(apple, index), Values(saved, index));
            var empty = Sheet(saved, 2);
            Assert.Equal("A1:C2", empty.Worksheet!.GetFirstChild<S.SheetDimension>()!.Reference!.Value);
            Assert.Equal(6, empty.Worksheet.Descendants<S.Cell>().Count());
            Assert.Empty(Values(saved, 2));
            // Apple positions this empty table at B2:D3; the adapter exports a table-relative A1:C2 grid.
            Assert.Equal("B2:D3", Sheet(apple, 2).Worksheet!.GetFirstChild<S.SheetDimension>()!.Reference!.Value);
            Assert.Equal(6, Sheet(apple, 2).Worksheet!.Descendants<S.Cell>().Count());
        }
        output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        Assert.Equal(3, reopened.Sheets.Count);
        Assert.Equal(20.75, reopened.Sheets[0].CellAt(2, 2).GetValue<double>());
        Assert.Equal("=literal", reopened.Sheets[0].CellAt(3, 4).GetValue<string>());
        Assert.Null(reopened.Sheets[0].GetFormulaText(3, 4));
    }

    [Fact]
    public void Tables_snapshot_input_and_failed_additions_preserve_collections() {
        var data = IWorkTableData.Create(2, 2).SetText(1, 1, "original");
        var document = IWorkNumbersDocument.Create(); var sheet = document.AddSheet("First");
        var table = sheet.AddTable(data, "Values"); byte[] before = document.SaveBytes();
        data.SetText(1, 1, "changed").SetBoolean(2, 2, true);
        Assert.Equal(before, document.SaveBytes()); Assert.Equal("original", Assert.Single(table.Cells).Value);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkTableCell>)table.Cells).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkNumbersTable>)sheet.Tables).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkNumbersSheetBuilder>)document.Sheets).Clear());
        Assert.Throws<ArgumentException>(() => document.AddSheet("FIRST"));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(data, "VALUES"));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddTable(data, "other", columnWidthPoints: float.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddTable(data, "other", rowHeightPoints: float.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddTable(data, "other", leftPoints: float.PositiveInfinity));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(data, "bad\nname"));
        Assert.Equal(before, document.SaveBytes());
        document.AddSheet("Second").AddTable(data, "Values"); // Names are unique within their own sheet.
        Assert.Equal(2, IWorkSourceDocument.Open(document.SaveBytes()).ReadNumbers().Sheets.Count);
    }

    [Theory]
    [InlineData("sheets")]
    [InlineData("tables")]
    [InlineData("cells")]
    [InlineData("text")]
    [InlineData("package")]
    public void Operation_budgets_reject_before_mutating_a_stream_or_existing_file(string boundary) {
        var document = Example(); var limits = new IWorkNumbersWriteOptions();
        switch (boundary) {
            case "sheets": limits.MaximumSheets = 1; break;
            case "tables": limits.MaximumTables = 2; break;
            case "cells": limits.MaximumTableCells = 25; break;
            case "text": limits.MaximumTextCharacters = 20; break;
            case "package": limits.MaximumPackageBytes = 200; break;
        }
        byte[] sentinel = { 9, 8, 7 }; using var stream = new MemoryStream(); stream.Write(sentinel);
        Assert.Throws<InvalidDataException>(() => document.Save(stream, limits)); Assert.Equal(sentinel, stream.ToArray());
        string directory = TemporaryFolder();
        try {
            string path = Path.Combine(directory, "existing.numbers"); File.WriteAllBytes(path, sentinel);
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: limits));
            Assert.Equal(sentinel, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public void Limits_count_UTF16_names_and_strings_and_all_unset_grid_positions() {
        var document = IWorkNumbersDocument.Create();
        document.AddSheet("S😀").AddTable(IWorkTableData.Create(2, 3).SetText(1, 1, "x😀"), "T");
        Assert.Throws<InvalidDataException>(() => document.SaveBytes(new() { MaximumTextCharacters = 6 }));
        Assert.Throws<InvalidDataException>(() => document.SaveBytes(new() { MaximumTableCells = 5 }));
        var table = IWorkSourceDocument.Open(document.SaveBytes(new() { MaximumTextCharacters = 7, MaximumTableCells = 6 }))
            .ReadNumbers().Sheets[0].Tables[0];
        Assert.Equal("x😀", Assert.Single(table.Cells).Value);
        Assert.Equal(2, table.RowCount); Assert.Equal(3, table.ColumnCount);
    }

    [Fact]
    public void Saves_obey_stream_position_conflicts_cancellation_and_caller_ownership() {
        var document = Example(); byte[] bytes = document.SaveBytes();
        using var stream = new MemoryStream(); stream.WriteByte(9); document.Save(stream);
        Assert.True(stream.CanWrite); Assert.Equal(bytes, stream.ToArray().Skip(1));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel(); byte[] before = stream.ToArray();
        Assert.Throws<OperationCanceledException>(() => document.Save(stream, cancellationToken: cancelled.Token));
        Assert.Equal(before, stream.ToArray());
        string directory = TemporaryFolder();
        try {
            string path = Path.Combine(directory, "created.numbers"); document.Save(path);
            Assert.Throws<IOException>(() => document.Save(path)); Assert.Equal(bytes, File.ReadAllBytes(path));
            Assert.Throws<OperationCanceledException>(() => document.Save(path, overwrite: true, cancellationToken: cancelled.Token));
            Assert.Equal(bytes, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, recursive: true); }
        var large = IWorkNumbersDocument.Create();
        large.AddSheet("Long").AddTable(IWorkTableData.Create(1, 1).SetText(1, 1, new string('a', 150_000)));
        using var copying = new CancellationTokenSource(); using var destination = new CancelAfterWriteStream(copying);
        Assert.Throws<OperationCanceledException>(() => large.Save(destination, cancellationToken: copying.Token));
        Assert.Equal(65_536, destination.Length); Assert.True(destination.CanWrite);
    }

    [Fact]
    public void Invalid_names_empty_documents_and_options_fail_before_publication() {
        var document = IWorkNumbersDocument.Create();
        Assert.Throws<InvalidOperationException>(() => document.SaveBytes());
        Assert.Throws<ArgumentNullException>(() => document.AddSheet(null!));
        Assert.Throws<ArgumentException>(() => document.AddSheet(" "));
        Assert.Throws<ArgumentException>(() => document.AddSheet("bad\tname"));
        Assert.Throws<ArgumentException>(() => document.AddSheet(new string('x', 129)));
        Assert.Throws<EncoderFallbackException>(() => document.AddSheet("invalid\uD800"));
        Assert.Empty(document.Sheets);
        var sheet = document.AddSheet("valid");
        Assert.Throws<InvalidOperationException>(() => document.SaveBytes());
        Assert.Throws<ArgumentNullException>(() => sheet.AddTable(null!));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(IWorkTableData.Create(1, 1), color: "invalid"));
        Assert.Empty(sheet.Tables);
        sheet.AddTable(IWorkTableData.Create(1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumSheets = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTables = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTableCells = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTextCharacters = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumPackageBytes = 0 }));
    }

    internal static IWorkNumbersDocument Example() {
        var document = IWorkNumbersDocument.Create();
        var data = IWorkTableData.Create(4, 4).SetText(1, 1, "Item").SetText(1, 2, "Value").SetText(1, 3, "Flag").SetText(1, 4, "Notes")
            .SetText(2, 1, "Apples").SetNumber(2, 2, 12.5).SetBoolean(2, 3, true).SetText(2, 4, "Zażółć 😀")
            .SetText(3, 1, "Negative").SetNumber(3, 2, -0.125).SetBoolean(3, 3, false).SetText(3, 4, "=literal")
            .SetText(4, 1, "Line1\nLine2").SetNumber(4, 2, 12345.6789).SetBoolean(4, 3, true);
        var first = document.AddSheet("Budget");
        first.AddTable(data, "Data", 40, 80, 130, 36, "Arial", 14, "003366");
        first.AddTable(IWorkTableData.Create(2, 2).SetText(1, 1, "Totals").SetNumber(1, 2, 99).SetBoolean(2, 1, false),
            "Totals", 40, 300, 130, 36);
        document.AddSheet("Empty grid").AddTable(IWorkTableData.Create(2, 3), "Empty", 30, 40, 80, 28);
        return document;
    }

    private static void VerifyTables(IWorkNumbersProjection projection, double editedValue) {
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal(new[] { "Budget", "Empty grid" }, projection.Sheets.Select(sheet => sheet.Name));
        Assert.Equal(new[] { "Data", "Totals", "Empty" }, projection.Sheets.SelectMany(sheet => sheet.Tables).Select(table => table.Name));
        Assert.All(projection.Sheets, sheet => Assert.Empty(sheet.TextBoxes));
        IWorkTable data = projection.Sheets[0].Tables[0];
        Assert.Equal(4, data.RowCount); Assert.Equal(4, data.ColumnCount);
        Assert.Equal(editedValue, data.GetCell(2, 2)!.Value);
        Assert.Equal(-0.125, data.GetCell(3, 2)!.Value); Assert.Equal(12345.6789, data.GetCell(4, 2)!.Value);
        Assert.Equal(true, data.GetCell(2, 3)!.Value); Assert.Equal(false, data.GetCell(3, 3)!.Value);
        Assert.Equal("Zażółć 😀", data.GetCell(2, 4)!.Value); Assert.Equal("Line1\nLine2", data.GetCell(4, 1)!.Value);
        Assert.Equal("=literal", data.GetCell(3, 4)!.Value); Assert.Equal(IWorkCellKind.Text, data.GetCell(3, 4)!.Kind);
        Assert.Null(data.GetCell(4, 4));
        IWorkTable totals = projection.Sheets[0].Tables[1];
        Assert.Equal(99d, totals.GetCell(1, 2)!.Value); Assert.Equal(false, totals.GetCell(2, 1)!.Value);
        IWorkTable empty = Assert.Single(projection.Sheets[1].Tables);
        Assert.Equal(2, empty.RowCount); Assert.Equal(3, empty.ColumnCount); Assert.Empty(empty.Cells);
    }

    private static WorksheetPart Sheet(SpreadsheetDocument document, int index) =>
        (WorksheetPart)document.WorkbookPart!.GetPartById(document.WorkbookPart.Workbook!.Sheets!.Elements<S.Sheet>().ElementAt(index).Id!);

    private static KeyValuePair<string, object>[] Values(SpreadsheetDocument document, int index) =>
        Sheet(document, index).Worksheet!.Descendants<S.Cell>().Where(cell => cell.CellValue is not null || cell.InlineString is not null)
            .Select(cell => new KeyValuePair<string, object>(cell.CellReference!.Value!, Value(document, cell)))
            .OrderBy(pair => pair.Key, StringComparer.Ordinal).ToArray();

    private static object Value(SpreadsheetDocument document, S.Cell cell) {
        if (cell.DataType?.Value == S.CellValues.InlineString) return cell.InlineString!.InnerText;
        if (cell.DataType?.Value == S.CellValues.Boolean) return cell.CellValue!.Text == "1";
        if (cell.DataType?.Value == S.CellValues.SharedString)
            return document.WorkbookPart!.SharedStringTablePart!.SharedStringTable!
                .Elements<S.SharedStringItem>().ElementAt(int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)).InnerText;
        if (cell.DataType?.Value == S.CellValues.String) return cell.CellValue!.Text;
        return double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
    private static string TemporaryFolder() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-NumbersWriting-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path); return path;
    }
    private sealed class CancelAfterWriteStream(CancellationTokenSource cancellation) : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) { base.Write(buffer, offset, count); cancellation.Cancel(); }
    }
}
