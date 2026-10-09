using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.Data;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("DataSet", false)]
        [InlineData("DataSet", true)]
        [InlineData("DataTable", false)]
        [InlineData("DataTable", true)]
        [InlineData("ParallelDataTable", false)]
        [InlineData("ParallelDataTable", true)]
        [InlineData("Projected", false)]
        [InlineData("Projected", true)]
        [InlineData("Objects", false)]
        [InlineData("Objects", true)]
        [InlineData("Dictionary", false)]
        [InlineData("Dictionary", true)]
        [InlineData("ReadOnlyDictionary", false)]
        [InlineData("ReadOnlyDictionary", true)]
        [InlineData("LegacyDictionary", false)]
        [InlineData("LegacyDictionary", true)]
        public void TabularImport_StyledMissingCells_RemainBlankAndKeepColumnFormat(string source, bool standardWriter) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet;
            if (source is "DataSet" or "DataTable" or "ParallelDataTable") {
                var table = new DataTable("Data");
                table.Columns.Add("Id", typeof(int));
                table.Columns.Add("Amount", typeof(double));
                table.Columns.Add("Text", typeof(string));
                table.Rows.Add(1, DBNull.Value, string.Empty);
                table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
                if (source == "DataSet") {
                    var dataSet = new DataSet();
                    dataSet.Tables.Add(table);
                    document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);
                    sheet = document.Sheets[0];
                } else {
                    sheet = document.AddWorksheet("Data");
                    sheet.InsertDataTable(table, mode: source == "ParallelDataTable" ? ExcelExecutionMode.Parallel : ExcelExecutionMode.Sequential);
                }
            } else {
                sheet = document.AddWorksheet("Data");
                if (source == "Projected") {
                    sheet.InsertObjects(new[] { 1, 2 },
                        ("Id", (Func<int, object?>)(id => id == 1 ? 1 : null)),
                        ("Amount", _ => null),
                        ("Text", id => id == 1 ? string.Empty : null));
                } else if (source == "Objects") {
                    sheet.InsertObjects(new[] {
                        new MissingStyleRow { Id = 1, Text = string.Empty },
                        new MissingStyleRow()
                    });
                } else if (source == "ReadOnlyDictionary") {
                    sheet.InsertObjects(new IReadOnlyDictionary<string, object?>[] {
                        new System.Collections.ObjectModel.ReadOnlyDictionary<string, object?>(new Dictionary<string, object?> { ["Id"] = 1, ["Amount"] = null, ["Text"] = string.Empty }),
                        new System.Collections.ObjectModel.ReadOnlyDictionary<string, object?>(new Dictionary<string, object?> { ["Id"] = null, ["Amount"] = DBNull.Value, ["Text"] = null })
                    });
                } else if (source == "LegacyDictionary") {
                    sheet.InsertObjects(new System.Collections.IDictionary[] {
                        new System.Collections.Specialized.OrderedDictionary { ["Id"] = 1, ["Amount"] = null, ["Text"] = string.Empty },
                        new System.Collections.Specialized.OrderedDictionary { ["Id"] = null, ["Amount"] = DBNull.Value, ["Text"] = null }
                    });
                } else {
                    sheet.InsertObjects(new[] {
                        new Dictionary<string, object?> { ["Id"] = 1, ["Amount"] = null, ["Text"] = string.Empty },
                        new Dictionary<string, object?> { ["Id"] = null, ["Amount"] = DBNull.Value, ["Text"] = null }
                    });
                }
            }
            if (source == "DataSet") {
                Assert.True(Assert.Single(sheet.ApplyColumnFormatPlan(new ExcelColumnFormatPlan().AddFormat("Amount", "0.0000"))).Applied);
            } else {
                sheet.ColumnStyleByHeader("Amount").NumberFormat("0.0000");
            }

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsx, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });

            using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            foreach (string reference in new[] { "B2", "B3" }) {
                Assert.True(cells.TryGetValue(reference, out Cell? cell), "Formatted blank cell " + reference + " must be retained.");
                Assert.Null(cell!.CellValue);
                Assert.Null(cell.CellFormula);
                Assert.Null(cell.InlineString);
                Assert.Null(cell.DataType);
                Assert.NotNull(cell.StyleIndex);
            }
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Equal("0.0000", loaded.Sheets[0].GetCellStyle(2, 2).NumberFormatCode);
            Assert.Equal("0.0000", loaded.Sheets[0].GetCellStyle(3, 2).NumberFormatCode);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
            Assert.Equal(3, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(1));
            Assert.False(reader.IsDBNull(2));
            Assert.Equal(string.Empty, reader.GetString(2));
            Assert.True(reader.Read());
            Assert.All(Enumerable.Range(0, 3), ordinal => Assert.True(reader.IsDBNull(ordinal)));
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx, false, false, false)]
        [InlineData(ExcelFileFormat.Xlsx, true, false, false)]
        [InlineData(ExcelFileFormat.Xlsx, false, true, false)]
        [InlineData(ExcelFileFormat.Xlsx, true, true, false)]
        [InlineData(ExcelFileFormat.Xls, false, false, false)]
        [InlineData(ExcelFileFormat.Xls, true, false, false)]
        [InlineData(ExcelFileFormat.Xls, false, true, false)]
        [InlineData(ExcelFileFormat.Xls, true, true, false)]
        [InlineData(ExcelFileFormat.Xlsb, false, false, false)]
        [InlineData(ExcelFileFormat.Xlsb, true, false, false)]
        [InlineData(ExcelFileFormat.Xlsb, false, true, false)]
        [InlineData(ExcelFileFormat.Xlsb, true, true, false)]
        [InlineData(ExcelFileFormat.Xlsb, false, false, true)]
        [InlineData(ExcelFileFormat.Xlsb, false, true, true)]
        public void TabularImport_ExplicitCellNullCompatibility_IsEmptyTextAcrossFormats(ExcelFileFormat format, bool standardWriter, bool batchWrite, bool useSharedStrings) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            var values = Enumerable.Range(1, 140).Select(row => (Row: row, Column: 1, Value: row == 2 ? DBNull.Value : row == 3 ? null! : (object)"Value"));
            if (batchWrite) {
                sheet.CellValues(values);
            } else {
                foreach (var value in values) sheet.CellValue(value.Row, value.Column, value.Value);
            }

            byte[] package = document.ToBytes(format, new ExcelSaveOptions {
                DisableFastPackageWriter = standardWriter,
                XlsbUseSharedStrings = format == ExcelFileFormat.Xlsb ? useSharedStrings : null
            });

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { HasHeaderRow = false });
            for (int row = 1; row <= 140; row++) {
                Assert.True(reader.Read());
                Assert.False(reader.IsDBNull(0));
                Assert.Equal(row is 2 or 3 ? string.Empty : "Value", reader.GetString(0));
            }
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData("DataTable", ExcelExecutionMode.Parallel)]
        [InlineData("DataTable", ExcelExecutionMode.Sequential)]
        [InlineData("Projected", ExcelExecutionMode.Sequential)]
        [InlineData("Projected", ExcelExecutionMode.Parallel)]
        [InlineData("SmallProjected", ExcelExecutionMode.Sequential)]
        [InlineData("Objects", ExcelExecutionMode.Sequential)]
        [InlineData("Objects", ExcelExecutionMode.Parallel)]
        [InlineData("SmallObjects", ExcelExecutionMode.Sequential)]
        [InlineData("CellValues", ExcelExecutionMode.Sequential)]
        [InlineData("CellValues", ExcelExecutionMode.Parallel)]
        public void TabularImport_PreparedOverwrite_ReplacesFormulaAndInlineTextAndPreservesStyles(string source, ExcelExecutionMode mode) {
            using ExcelDocument document = ExcelDocument.Create();
            document.Execution.Mode = mode;
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            sheet.CellFormula(2, 1, "123");
            sheet.CellValue(2, 2, "Previous");
            sheet.FormatCell(2, 1, "0.0000");
            sheet.FormatCell(2, 2, "@");
            Cell inline = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B2");
            inline.CellValue = null;
            inline.DataType = CellValues.InlineString;
            inline.InlineString = new InlineString(new Text("Previous"));
            if (source == "DataTable") {
                var table = new DataTable();
                table.Columns.Add("Formula", typeof(string));
                table.Columns.Add("Inline", typeof(string));
                table.Rows.Add(string.Empty, string.Empty);
                table.Rows.Add("Next", "Next");
                sheet.InsertDataTable(table, mode: mode);
            } else if (source is "Projected" or "SmallProjected") {
                sheet.InsertObjects(Enumerable.Range(1, source == "SmallProjected" ? 2 : 20),
                    ("Formula", (Func<int, object?>)(id => id == 1 ? string.Empty : "Next")),
                    ("Inline", id => id == 1 ? string.Empty : "Next"));
            } else if (source is "Objects" or "SmallObjects") {
                sheet.InsertObjects(Enumerable.Range(1, source == "SmallObjects" ? 2 : 20).Select(id => new OverwriteRow {
                    Formula = id == 1 ? string.Empty : "Next",
                    Inline = id == 1 ? string.Empty : "Next"
                }));
            } else {
                var cells = Enumerable.Range(2, 20).SelectMany(row => new[] {
                    (Row: row, Column: 1, Value: (object)(row == 2 ? string.Empty : "Next")),
                    (Row: row, Column: 2, Value: (object)(row == 2 ? string.Empty : "Next"))
                });
                sheet.CellValues(cells, mode: mode);
            }

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsx);

            using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
            var written = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            foreach (string reference in new[] { "A2", "B2" }) {
                Assert.Null(written[reference].CellFormula);
                Assert.Null(written[reference].InlineString);
            }
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Equal("0.0000", loaded.Sheets[0].GetCellStyle(2, 1).NumberFormatCode);
            Assert.Equal("@", loaded.Sheets[0].GetCellStyle(2, 2).NumberFormatCode);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { SheetName = "Data", HasHeaderRow = source != "CellValues", A1Range = source == "CellValues" ? "A2:B21" : null });
            Assert.True(reader.Read());
            Assert.False(reader.IsDBNull(0));
            Assert.False(reader.IsDBNull(1));
            Assert.Equal(string.Empty, reader.GetString(0));
            Assert.Equal(string.Empty, reader.GetString(1));
            Assert.True(reader.Read());
            Assert.Equal("Next", reader.GetString(0));
            Assert.Equal("Next", reader.GetString(1));
        }

        [Theory]
        [InlineData("DataTable")]
        [InlineData("Projected")]
        public void TabularImport_RejectedReplacement_PreservesTheExistingCell(string source) {
            using ExcelDocument document = ExcelDocument.Create();
            document.Execution.Mode = ExcelExecutionMode.Sequential;
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            sheet.CellFormula(2, 1, "123");
            sheet.FormatCell(2, 1, "0.0000");
            Cell existing = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
            string original = existing.OuterXml;
            string rejected = new string('a', 32768);

            if (source == "DataTable") {
                var table = new DataTable();
                table.Columns.Add("Value", typeof(string));
                table.Rows.Add(rejected);
                Assert.Throws<ArgumentException>(() => sheet.InsertDataTable(table, mode: ExcelExecutionMode.Sequential));
            } else {
                Assert.Throws<ArgumentException>(() => sheet.InsertObjects(new[] { 1, 2 },
                    ("Value", (Func<int, object?>)(_ => rejected))));
            }

            Assert.Equal(original, existing.OuterXml);
        }

        [Theory]
        [InlineData("DataSet", false)]
        [InlineData("DataSet", true)]
        [InlineData("DataTable", false)]
        [InlineData("DataTable", true)]
        [InlineData("ParallelDataTable", false)]
        [InlineData("ParallelDataTable", true)]
        [InlineData("Objects", false)]
        [InlineData("Objects", true)]
        public void TabularImport_HeaderlessStyledMissingRows_PreserveAllColumns(string source, bool standardWriter) {
            using ExcelDocument document = ExcelDocument.Create();
            if (source == "Objects") {
                document.AddWorksheet("Data").InsertObjects(new[] {
                    new HeaderlessDateStyleRow(), new HeaderlessDateStyleRow()
                }, includeHeaders: false);
            } else {
                var table = new DataTable("Data");
                table.Columns.Add("Id", typeof(int));
                table.Columns.Add("Date", typeof(DateTime));
                table.Columns.Add("Text", typeof(string));
                table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
                table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
                if (source == "DataSet") {
                    var dataSet = new DataSet();
                    dataSet.Tables.Add(table);
                    document.InsertDataSet(dataSet, createTables: false, includeHeaders: false, includeAutoFilter: false);
                } else {
                    document.AddWorksheet("Data").InsertDataTable(table, includeHeaders: false,
                        mode: source == "ParallelDataTable" ? ExcelExecutionMode.Parallel : ExcelExecutionMode.Sequential);
                }
            }

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsx, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });

            using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            for (int row = 1; row <= 2; row++) {
                foreach (string column in new[] { "A", "B", "C" }) {
                    string reference = column + row;
                    Assert.True(cells.TryGetValue(reference, out Cell? cell), "Declared column boundary or styled blank " + reference + " must be retained.");
                    Assert.Null(cell!.CellValue);
                    Assert.Null(cell.CellFormula);
                    Assert.Null(cell.InlineString);
                    Assert.Null(cell.DataType);
                }
                Assert.NotNull(cells["B" + row].StyleIndex);
            }
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Contains("yy", loaded.Sheets[0].GetCellStyle(1, 2).NumberFormatCode, StringComparison.OrdinalIgnoreCase);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(3, reader.FieldCount);
            for (int row = 0; row < 2; row++) {
                Assert.True(reader.Read());
                Assert.All(Enumerable.Range(0, 3), ordinal => Assert.True(reader.IsDBNull(ordinal)));
            }
            Assert.False(reader.Read());
        }

        private sealed class HeaderlessDateStyleRow {
            public int? Id { get; set; }
            public DateTime? Date { get; set; }
            public string? Text { get; set; }
        }

        private sealed class MissingStyleRow {
            public int? Id { get; set; }
            public double? Amount { get; set; }
            public string? Text { get; set; }
        }

        private sealed class OverwriteRow {
            public string Formula { get; set; } = string.Empty;
            public string Inline { get; set; } = string.Empty;
        }
    }
}
