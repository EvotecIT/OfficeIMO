using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(ExcelFileFormat.Xls, false, false)]
        [InlineData(ExcelFileFormat.Xls, false, true)]
        [InlineData(ExcelFileFormat.Xls, true, false)]
        [InlineData(ExcelFileFormat.Xls, true, true)]
        [InlineData(ExcelFileFormat.Xlsb, false, false)]
        [InlineData(ExcelFileFormat.Xlsb, false, true)]
        [InlineData(ExcelFileFormat.Xlsb, true, false)]
        [InlineData(ExcelFileFormat.Xlsb, true, true)]
        public void NativeBinarySave_ClearedTrailingCellsRetainOnlyIntentionalBlanks(ExcelFileFormat format, bool styled, bool clearAll) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Keep");
            sheet.CellValue(5, 4, 42);
            if (styled) sheet.FormatRange("D5:D5", "0.0000");
            sheet.ClearRange("D5:D5", clearAll ? ExcelClearOptions.All : ExcelClearOptions.Values);

            Cell cleared = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "D5");
            Assert.Null(cleared.CellValue);
            Assert.Null(cleared.DataType);
            bool retainStyle = styled && !clearAll;
            byte[] bytes = document.ToBytes(format);
            if (format == ExcelFileFormat.Xlsb) {
                using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(bytes, writable: false));
                ExcelSheet result = Assert.Single(reopened.Sheets);
                Cell[] cells = result.WorksheetPart.Worksheet.Descendants<Cell>().ToArray();
                Assert.Equal(retainStyle ? 2 : 1, cells.Length);
                Assert.True(result.TryGetCellText(1, 1, out string kept));
                Assert.Equal("Keep", kept);
                if (retainStyle) {
                    Cell blank = cells.Single(cell => cell.CellReference?.Value == "D5");
                    Assert.Null(blank.CellValue);
                    Assert.Null(blank.DataType);
                    Assert.Equal("0.0000", result.GetCellStyle(5, 4).NumberFormatCode);
                }
                // XLSB retains authored worksheet geometry independently of cells.
                Assert.Equal("A1:D5", result.WorksheetPart.Worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(bytes, new ExcelReadOptions { HasHeaderRow = false });
                Assert.Equal(retainStyle ? 4 : 1, reader.FieldCount);
                Assert.True(reader.Read());
                Assert.Equal("Keep", reader.GetString(0));
                int rowCount = 1;
                while (reader.Read()) {
                    Assert.All(Enumerable.Range(0, reader.FieldCount), ordinal => Assert.True(reader.IsDBNull(ordinal)));
                    rowCount++;
                }
                Assert.Equal(retainStyle ? 5 : 1, rowCount);
                return;
            }
            using MemoryStream stream = new MemoryStream(bytes, writable: false);
            using LegacyXlsLoadResult loaded = ExcelDocument.LoadLegacyXlsWithReport(stream);
            loaded.EnsureNoImportErrors();
            Assert.False(loaded.HasUnsupportedFeatures);
            LegacyXlsWorksheet worksheet = Assert.Single(loaded.Workbook.Worksheets);
            Assert.Equal(retainStyle ? "A1:D5" : "A1:A1", worksheet.DeclaredUsedRange!.UsedRangeA1);
            Assert.Equal(retainStyle ? 2 : 1, worksheet.Cells.Count);
            Assert.Equal("Keep", worksheet.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Value);
            if (retainStyle) {
                LegacyXlsCell blank = worksheet.Cells.Single(cell => cell.Row == 5 && cell.Column == 4);
                Assert.Equal(LegacyXlsCellValueKind.Blank, blank.Kind);
                Assert.NotEqual((ushort)0, loaded.Workbook.CellFormats[blank.StyleIndex].NumberFormatId);
            }
        }

        [Theory]
        [InlineData("String")]
        [InlineData("InlineString")]
        [InlineData("SharedString")]
        public void LegacyXls_NativeSave_PresentEmptyTextExtendsUsedRangeWithoutStyle(string storage) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Keep");
            sheet.CellValue(5, 4, string.Empty);
            Cell source = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "D5");
            source.RemoveAllChildren();
            source.DataType = storage == "InlineString" ? CellValues.InlineString
                : storage == "SharedString" ? CellValues.SharedString : CellValues.String;
            if (storage == "InlineString") {
                source.InlineString = new InlineString(new Text(string.Empty));
            } else if (storage == "SharedString") {
                SharedStringTablePart part = document.WorkbookPartRoot!.SharedStringTablePart
                    ?? document.WorkbookPartRoot.AddNewPart<SharedStringTablePart>();
                part.SharedStringTable ??= new SharedStringTable();
                int index = part.SharedStringTable.Elements<SharedStringItem>().Count();
                part.SharedStringTable.Append(new SharedStringItem(new Text(string.Empty)));
                source.CellValue = new CellValue(index.ToString(System.Globalization.CultureInfo.InvariantCulture));
            } else {
                source.CellValue = new CellValue(string.Empty);
            }
            byte[] xls = document.ToBytes(ExcelFileFormat.Xls);
            using LegacyXlsLoadResult loaded = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(xls, writable: false));
            loaded.EnsureNoImportErrors();
            LegacyXlsWorksheet worksheet = Assert.Single(loaded.Workbook.Worksheets);
            Assert.Equal("A1:D5", worksheet.DeclaredUsedRange!.UsedRangeA1);
            LegacyXlsCell emptyText = worksheet.Cells.Single(cell => cell.Row == 5 && cell.Column == 4);
            Assert.Equal(LegacyXlsCellValueKind.Text, emptyText.Kind);
            Assert.Equal(string.Empty, emptyText.Value);
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xls, false)]
        [InlineData(ExcelFileFormat.Xls, true)]
        [InlineData(ExcelFileFormat.Xlsb, false)]
        [InlineData(ExcelFileFormat.Xlsb, true)]
        public void NativeBinarySave_ImportedMissingValuesRemainDistinctFromEmptyText(ExcelFileFormat format, bool standardWriter) {
            DataTable table = new DataTable("Data");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Amount", typeof(double));
            table.Columns.Add("Text", typeof(string));
            table.Rows.Add(1, DBNull.Value, string.Empty);
            table.Rows.Add(2, DBNull.Value, "Next");
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.InsertDataTable(table);
            sheet.ColumnStyleByHeader("Amount").NumberFormat("0.0000");
            byte[] bytes = document.ToBytes(format, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(bytes);
            Assert.Equal(3, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            Assert.False(reader.IsDBNull(2));
            Assert.Equal(string.Empty, reader.GetString(2));
            Assert.True(reader.Read());
            Assert.Equal(2, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            Assert.Equal("Next", reader.GetString(2));
            Assert.False(reader.Read());
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(bytes, writable: false));
            Assert.Equal("0.0000", loaded.Sheets[0].GetCellStyle(2, 2).NumberFormatCode);
            Assert.Equal("0.0000", loaded.Sheets[0].GetCellStyle(3, 2).NumberFormatCode);
        }

        [Theory]
        [InlineData(false, ExcelFileFormat.Xls, false)]
        [InlineData(false, ExcelFileFormat.Xls, true)]
        [InlineData(true, ExcelFileFormat.Xls, false)]
        [InlineData(true, ExcelFileFormat.Xls, true)]
        [InlineData(false, ExcelFileFormat.Xlsb, false)]
        [InlineData(false, ExcelFileFormat.Xlsb, true)]
        [InlineData(true, ExcelFileFormat.Xlsb, false)]
        [InlineData(true, ExcelFileFormat.Xlsb, true)]
        public void NativeBinarySave_HeaderlessAllMissingImportsPreserveRowsAndColumns(bool dataSet, ExcelFileFormat format, bool standardWriter) {
            DataTable table = new DataTable("Data");
            table.Columns.Add("A", typeof(string));
            table.Columns.Add("B", typeof(string));
            table.Columns.Add("C", typeof(string));
            for (int row = 0; row < 65; row++) table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
            using ExcelDocument document = ExcelDocument.Create();
            if (dataSet) {
                DataSet set = new DataSet();
                set.Tables.Add(table);
                document.InsertDataSet(set, createTables: false, includeHeaders: false, includeAutoFilter: false);
            } else {
                document.AddWorksheet("Data").InsertDataTable(table, includeHeaders: false);
            }
            byte[] bytes = document.ToBytes(format, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(bytes, new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(3, reader.FieldCount);
            for (int row = 0; row < 65; row++) {
                Assert.True(reader.Read());
                Assert.All(Enumerable.Range(0, 3), ordinal => Assert.True(reader.IsDBNull(ordinal)));
            }
            Assert.False(reader.Read());
        }

        [Fact]
        public void LegacyXls_NativeSave_MissingReplacementKeepsDefaultPresenceUntilStylesAreCleared() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Replace");
            DataTable table = new DataTable();
            table.Columns.Add("A", typeof(string));
            table.Columns.Add("B", typeof(string));
            table.Columns.Add("C", typeof(string));
            table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
            sheet.InsertDataTable(table, includeHeaders: false, mode: ExcelExecutionMode.Parallel);
            sheet.ClearRange("A1:C1", ExcelClearOptions.Values);
            byte[] missing = document.ToBytes(ExcelFileFormat.Xls, new ExcelSaveOptions { DisableFastPackageWriter = true });
            using (ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(missing, new ExcelReadOptions { HasHeaderRow = false })) {
                Assert.Equal(3, reader.FieldCount);
                Assert.True(reader.Read());
                Assert.All(Enumerable.Range(0, 3), ordinal => Assert.True(reader.IsDBNull(ordinal)));
                Assert.False(reader.Read());
            }
            sheet.ClearRange("A1:C1", ExcelClearOptions.All);
            byte[] cleared = document.ToBytes(ExcelFileFormat.Xls, new ExcelSaveOptions { DisableFastPackageWriter = true });
            using ExcelWorkbookDataReader empty = ExcelDocument.OpenDataReader(cleared, new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(0, empty.FieldCount);
            Assert.False(empty.Read());
        }
    }
}
