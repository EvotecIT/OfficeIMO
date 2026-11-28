using System.Data;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void PerformanceReview_Utf8ObjectReaderDoesNotCoerceFormulaTextIntoTypedProperties(bool useCachedFormulaResult) {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
            })) {
                var sheet = document.AddSheet("Formula");
                sheet.Cell(2, 1).SetValue("Id");
                sheet.Cell(2, 2).SetValue("Active");
                sheet.Cell(2, 3).SetValue("Expression");
                sheet.Cell(3, 1).SetFormula("1");
                sheet.Cell(3, 2).SetFormula("TRUE");
                sheet.Cell(3, 3).SetFormula("TRUE");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray(), new ExcelReaderOptions {
                UseCachedFormulaResult = useCachedFormulaResult
            });
            FormulaProjection row = reader.GetSheet("Formula")
                .ReadObjectsStream<FormulaProjection>("A2:C3")
                .Single();

            Assert.Equal(0, row.Id);
            Assert.False(row.Active);
            Assert.Equal("TRUE", row.Expression);
        }

        [Fact]
        public void PerformanceReview_XmlDataReaderDoesNotReusePrimitiveValuesForMissingCells() {
            var expectedDate = new DateTime(2026, 7, 14);
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
            })) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Id");
                sheet.Cell(1, 2).SetValue("Created");
                sheet.Cell(1, 3).SetValue("Active");
                sheet.Cell(1, 4).SetValue("Note");
                sheet.Cell(2, 1).SetValue(42);
                sheet.Cell(2, 2).SetValue(expectedDate);
                sheet.Cell(2, 3).SetValue(true);
                sheet.Cell(2, 4).SetValue("Complete");
                sheet.Cell(3, 4).SetValue("Missing typed values");
            }

            byte[] utf16WorksheetPackage = RewriteFirstWorksheetAsUtf16(memory.ToArray());
            using var reader = ExcelDocumentReader.Open(utf16WorksheetPackage);
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:D3", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal(42, dataReader.GetInt32(0));
            Assert.Equal(expectedDate, dataReader.GetDateTime(1));
            Assert.True(dataReader.GetBoolean(2));

            Assert.True(dataReader.Read());
            Assert.Equal(DBNull.Value, dataReader.GetValue(0));
            Assert.Equal(DBNull.Value, dataReader.GetValue(1));
            Assert.Equal(DBNull.Value, dataReader.GetValue(2));
            Assert.Equal("Missing typed values", dataReader.GetString(3));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void PerformanceReview_DashboardChartOverDataTableTableKeepsTheTableWorksheetData(bool includeCachedData) {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Score", typeof(double));
            table.Columns.Add("Note", typeof(string));
            for (int i = 0; i < 40; i++) table.Rows.Add(i % 2 == 0 ? "North" : "South", i + 0.5, "n" + i);

            byte[] Save(bool addChart) {
                using var stream = new MemoryStream();
                using (var document = ExcelDocument.Create(stream, new ExcelCreateOptions {
                    PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
                })) {
                    var sheet = document.AddSheet("Data");
                    sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
                    if (addChart) {
                        OfficeIMO.Excel.Reports.ExcelReportExtensions.AddDashboardChart(sheet, new OfficeIMO.Excel.Reports.ExcelDashboardChartOptions {
                            Range = "A1:B41", Row = 2, Column = 6, Title = "Scores", IncludeCachedData = includeCachedData
                        });
                    }
                }

                return stream.ToArray();
            }

            string SheetData(byte[] package) {
                using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
                using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open());
                string xml = reader.ReadToEnd();
                int start = xml.IndexOf("<sheetData", StringComparison.Ordinal);
                return xml.Substring(start, xml.IndexOf("</sheetData>", StringComparison.Ordinal) - start);
            }

            byte[] withChart = Save(addChart: true);

            Assert.Equal(SheetData(Save(addChart: false)), SheetData(withChart));
            using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(withChart), false);
            var chartPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            Assert.Equal(includeCachedData ? 40 : 0, chartPart.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumericPoint>().Count());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Theory]
        [InlineData("range")]
        [InlineData("tables")]
        public void PerformanceReview_TableMetadataReadsKeepTheDataTableWorksheetData(string read) {
            var table = new DataTable();
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Enabled", typeof(bool));
            table.Columns.Add("Created", typeof(DateTime));
            table.Columns.Add("Score", typeof(decimal));
            for (int i = 1; i <= 30; i++) table.Rows.Add(i, "n" + i, i % 2 == 0, new DateTime(2024, 1, 1).AddMinutes(i), i * 1.25m);

            string SheetData(bool readMetadata) {
                using var stream = new MemoryStream();
                using (var document = ExcelDocument.Create(stream, new ExcelCreateOptions {
                    PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
                })) {
                    var sheet = document.AddSheet("Data");
                    var created = sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
                    if (readMetadata && read == "range") Assert.Equal("A1:E31", created.Range);
                    if (readMetadata && read == "tables") Assert.Equal("A1:E31", Assert.Single(sheet.GetTables()).Range);
                }

                using var archive = new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read);
                using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open());
                string xml = reader.ReadToEnd();
                int start = xml.IndexOf("<sheetData", StringComparison.Ordinal);
                return xml.Substring(start, xml.IndexOf("</sheetData>", StringComparison.Ordinal) - start);
            }

            Assert.Equal(SheetData(readMetadata: false), SheetData(readMetadata: true));
        }

        [Theory]
        [InlineData(ExcelTableStyle.TableStyleMedium9, true)]
        [InlineData(ExcelTableStyle.None, false)]
        public void PerformanceReview_PendingDataTableReportsTheTableTheWorkbookLoads(ExcelTableStyle style, bool includeAutoFilter) {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(object));
            table.Columns.Add("Score", typeof(double));
            table.Rows.Add("NA", 1.5);
            table.Rows.Add("EU", 2.5);

            static string Describe(ExcelTableInfo info) => string.Join("|", info.Name, info.DisplayName, info.Range, info.SheetName, info.SheetIndex,
                info.StyleName, info.HasHeaderRow, info.TotalsRowShown, info.HasAutoFilter, info.ShowFirstColumn, info.ShowLastColumn,
                info.ShowRowStripes, info.ShowColumnStripes, string.Join(",", info.Columns.Select(column => column.Index + ":" + column.Name + ":" + column.TotalsRowFunction)));

            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddSheet("Data");
            sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", Style = style, IncludeAutoFilter = includeAutoFilter, AutoFit = false }, row: 1, column: 1);

            string pending = Describe(Assert.Single(sheet.GetTables()));

            // The workbook-wide listing loads the rows and reads the table part.
            Assert.Equal(Describe(Assert.Single(document.GetTables())), pending);
            Assert.Equal(pending, Describe(Assert.Single(sheet.GetTables())));
        }

        [Fact]
        public void PerformanceReview_ReportEditsOverAPendingDataTableMatchTheLoadedWorksheet() {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(object));
            table.Columns.Add("Score", typeof(object));
            table.Columns.Add("Created", typeof(object));
            table.Columns.Add("Count", typeof(object));
            for (int i = 1; i <= 30; i++) table.Rows.Add(i % 2 == 0 ? "NA" : "EU", i * 1.137, new DateTime(2024, 1, 1).AddMinutes(i), i % 7);

            // A cell beside the table before the edits makes the header lookups load the rows into the worksheet;
            // written last, the edits run while the rows are still pending.
            byte[] Save(bool loadRowsFirst) {
                using var stream = new MemoryStream();
                using (var document = ExcelDocument.Create(stream)) {
                    var sheet = document.AddSheet("Data");
                    sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
                    if (loadRowsFirst) sheet.Cell(1, 12).SetValue("note");
                    Assert.True(sheet.TryGetColumnRangeByHeader("Score", null, 0, false, true, out string score));
                    Assert.Equal("B2:B31", score);
                    sheet.Range(score).SetNumberFormat("#,##0.000");
                    Assert.True(sheet.TryGetColumnRangeByHeader("created", null, 0, true, true, out string created));
                    Assert.Equal("C1:C31", created);
                    sheet.Range("C2:C31").SetNumberFormat("yyyy-mm-dd hh:mm");
                    sheet.Range(score).ConditionalFormat().CellIs(ExcelConditionalFormattingOperator.GreaterThan, "10", fillColor: OfficeColor.Parse("#F4CCCC"));
                    Assert.Single(sheet.GetTables());
                    sheet.AddChart(OfficeChartType.ColumnClustered, "A1:B31", new ExcelChartOptions { Row = 2, Column = 6 });
                    sheet.AddPivotTable("A1:D31", "F20", new ExcelPivotTableOptions {
                        Name = "Summary",
                        RowFields = new List<string> { "Region" },
                        DataFields = new List<ExcelPivotDataField> {
                            new ExcelPivotDataField("Count", ExcelPivotDataFunction.Sum, "Total", numberFormat: "#,##0"),
                            new ExcelPivotDataField("Score", ExcelPivotDataFunction.Average, "Average", numberFormat: "0.0000")
                        }
                    });
                    if (!loadRowsFirst) sheet.Cell(1, 12).SetValue("note");
                    document.Save(stream);
                }

                return stream.ToArray();
            }

            // Each cell as reference, text and number format; shared and inline strings read the same.
            List<string> Cells(byte[] package) {
                using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package), false);
                var workbookPart = spreadsheet.WorkbookPart!;
                var strings = workbookPart.SharedStringTablePart?.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>()
                    .Select(item => item.InnerText).ToArray() ?? Array.Empty<string>();
                var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet;
                var formats = stylesheet.NumberingFormats?.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>()
                    .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode!.Value!) ?? new Dictionary<uint, string>();
                var cellFormats = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ToArray();
                var result = new List<string>();
                foreach (var cell in workbookPart.WorksheetParts.Single().Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                    string text = cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                        ? strings[int.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture)]
                        : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
                    // Compare numbers as values: the two writers may spell the same double with different digit counts.
                    if (cell.DataType == null && double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number)) {
                        text = number.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                    }

                    uint formatId = cellFormats[(int)(cell.StyleIndex?.Value ?? 0U)].NumberFormatId?.Value ?? 0U;
                    result.Add(cell.CellReference + "=" + text + "|" + (formats.TryGetValue(formatId, out string? code) ? code : formatId.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }

                // Conditional formats and pivot measures refer to workbook styles by id.
                var worksheetPart = workbookPart.WorksheetParts.Single();
                foreach (var rule in worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule>()) {
                    var dxf = rule.FormatId == null ? null : stylesheet.DifferentialFormats?.Elements<DocumentFormat.OpenXml.Spreadsheet.DifferentialFormat>().ElementAtOrDefault((int)rule.FormatId.Value);
                    result.Add("rule " + rule.Type?.InnerText + "=" + (dxf?.InnerXml ?? "none"));
                }

                foreach (var dataField in worksheetPart.PivotTableParts.Single().PivotTableDefinition!.DataFields!.Elements<DocumentFormat.OpenXml.Spreadsheet.DataField>()) {
                    uint id = dataField.NumberFormatId?.Value ?? 0U;
                    result.Add("measure " + dataField.Name + "=" + (formats.TryGetValue(id, out string? code) ? code : id.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }

                return result;
            }

            byte[] pending = Save(loadRowsFirst: false);
            List<string> cells = Cells(pending);

            Assert.Equal(Cells(Save(loadRowsFirst: true)), cells);
            Assert.Contains("B2=1.137|#,##0.000", cells);
            Assert.Contains("C1=Created|0", cells);
            Assert.Contains(cells, cell => cell.StartsWith("C31=45292.02", StringComparison.Ordinal) && cell.EndsWith("|yyyy-mm-dd hh:mm", StringComparison.Ordinal));
            Assert.Contains(cells, cell => cell.StartsWith("D2=", StringComparison.Ordinal) && cell.EndsWith("|0", StringComparison.Ordinal));
            Assert.Contains(cells, cell => cell.StartsWith("rule cellIs=", StringComparison.Ordinal) && cell.Contains("FFF4CCCC"));
            Assert.Contains("measure Average=0.0000", cells);
            using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(pending), false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Assert.Single(worksheetPart.PivotTableParts);
            Assert.Single(worksheetPart.DrawingsPart!.ChartParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Theory]
        [InlineData("range")]
        [InlineData("ranges")]
        [InlineData("table")]
        public void PerformanceReview_RangeChartCacheFromDeferredRowsMatchesTheWrittenCells(string source) {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(object));
            table.Columns.Add("Score", typeof(object));
            table.Columns.Add("Count", typeof(object));
            object?[][] rows = {
                new object?[] { "NA", 1.0 / 3.0, 7 },
                new object?[] { "EU & UK", 12.5m, 3L },
                new object?[] { 2024, 1e21, (byte)4 },
                new object?[] { null, null, 0.1f },
                new object?[] { "LATAM", "42", -5 }
            };
            foreach (object?[] row in rows) table.Rows.Add(row);

            string ChartXml(bool writeRowsFirst) {
                using var stream = new MemoryStream();
                using (var document = ExcelDocument.Create(stream)) {
                    var sheet = document.AddSheet("Data");
                    sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
                    if (writeRowsFirst) {
                        sheet.Cell(1, 10).SetValue("note");
                        // A cell read loads the rows into the worksheet, so the chart reads the written cells.
                        Assert.True(sheet.TryGetCellText(2, 1, out _));
                    }
                    if (source == "ranges") {
                        sheet.AddChartFromRanges(OfficeChartType.ColumnClustered, new[] {
                            new ExcelChartSeriesRange("Score", "A2:A6", "B2:B6"),
                            new ExcelChartSeriesRange("Count", "A2:A6", "C2:C6")
                        }, new ExcelChartOptions { Row = 2, Column = 6 });
                    } else {
                        sheet.AddChart(OfficeChartType.ColumnClustered, source == "table" ? "Scores" : "A1:C6", new ExcelChartOptions { Row = 2, Column = 6 });
                    }
                    // Reading the cache must leave the rows for the direct writer.
                    Assert.Equal(!writeRowsFirst, document.HasUnmaterializedDirectDataSetRows);
                    if (!writeRowsFirst) sheet.Cell(1, 10).SetValue("note");
                    document.Save(stream);
                }

                using var archive = new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read);
                var entry = archive.Entries.Single(item => item.FullName.Contains("/charts/chart") && item.FullName.EndsWith(".xml", StringComparison.Ordinal)
                    && !item.FullName.Contains("/_rels/"));
                using var reader = new StreamReader(entry.Open());
                string xml = reader.ReadToEnd();
                return string.Join("|", System.Text.RegularExpressions.Regex.Matches(xml, "<c:(strCache|numCache)>.*?</c:\\1>")
                    .Cast<System.Text.RegularExpressions.Match>().Select(match => match.Value));
            }

            string deferred = ChartXml(writeRowsFirst: false);
            Assert.Contains("EU &amp; UK", deferred);
            Assert.Equal(ChartXml(writeRowsFirst: true), deferred);
        }

        [Fact]
        public void PerformanceReview_ChartedTableSavesItsRowsWithoutLoadingThemIntoTheWorksheet() {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(object));
            table.Columns.Add("Score", typeof(object));
            table.Rows.Add("NA", 1.5);
            table.Rows.Add("EU", 2.5);

            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                var sheet = document.AddSheet("Data");
                sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
                sheet.AddChart(OfficeChartType.ColumnClustered, "Scores", new ExcelChartOptions { Row = 2, Column = 4 });
                Assert.True(document.HasWorksheets);
                document.Save(stream);
                // Neither the worksheet check nor the save's look for a drawing may load the rows into the worksheet.
                Assert.True(document.HasUnmaterializedDirectDataSetRows);
            }

            using var saved = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            Assert.True(saved["Data"].TryGetCellText(3, 1, out string region));
            Assert.Equal("EU", region);
            Assert.Single(saved["Data"].Charts);
        }

        [Fact]
        public void PerformanceReview_SheetViewEditsAfterFreezingADataTableKeepBothSettings() {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(object));
            table.Columns.Add("Sales", typeof(object));
            table.Rows.Add("NA", 100);
            table.Rows.Add("EMEA", 200);

            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                var sheet = document.AddSheet("Data");
                var created = sheet.AddTable(table, new ExcelTableOptions { Name = "Sales", AutoFit = false }, row: 1, column: 1);
                Assert.Equal("A1:B3", created.Range);
                sheet.View.Freeze(1);
                sheet.View.ShowGridlines = false;
                document.Save(stream);
            }

            using var loaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            var view = loaded.Sheets[0].View;
            Assert.False(view.ShowGridlines);
            Assert.Equal(1, view.FrozenRows);
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string last));
            Assert.Equal("EMEA", last);
        }

        [Fact]
        public void PerformanceReview_PackageEditsAfterReadingDataTableCellsReachTheSavedWorksheet() {
            var table = new DataTable();
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(object));
            for (int i = 1; i <= 10; i++) table.Rows.Add(i, "n" + i);

            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                var sheet = document.AddSheet("Data");
                sheet.AddTable(table, new ExcelTableOptions { Name = "Names", AutoFit = false }, row: 1, column: 1);
                Assert.True(sheet.TryGetCellText(3, 2, out _));
                var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
                var cell = worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single(item => item.CellReference?.Value == "B3");
                cell.InlineString = null;
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("edited");
                document.Save(stream);
            }

            using var loaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 2, out string text));
            Assert.Equal("edited", text);
            Assert.True(loaded.Sheets[0].TryGetCellText(11, 2, out string last));
            Assert.Equal("n10", last);
        }

        [Fact]
        public void PerformanceReview_EditsAfterReadingDataTableCellsReachTheSavedWorksheet() {
            var table = new DataTable();
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(object));
            for (int i = 1; i <= 30; i++) table.Rows.Add(i, "n" + i);

            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                var sheet = document.AddSheet("Data");
                sheet.AddTable(table, new ExcelTableOptions { Name = "Names", AutoFit = false }, row: 1, column: 1);
                Assert.True(sheet.TryGetCellText(3, 2, out string text));
                Assert.Equal("n2", text);
                sheet.Cell(1, 1).SetBold(true);
                sheet.Cell(5, 2).SetValue("edited");
                sheet.AutoFitColumnsFor(new[] { 1, 2 });
                document.Save(stream);
            }

            using var loaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            var saved = loaded.Sheets[0];
            Assert.True(saved.TryGetCellText(5, 2, out string edited));
            Assert.Equal("edited", edited);
            Assert.True(saved.TryGetCellText(31, 2, out string last));
            Assert.Equal("n30", last);
            Assert.True(saved.Cell(1, 1).GetStyle().Bold);
        }

        [Fact]
        public void CellValueSnapshots_ColumnReadMatchesSingleCellReadsOnLoadedAndImportedSheets() {
            byte[] package;
            using (var stream = new MemoryStream()) {
                using (var document = ExcelDocument.Create(stream, new ExcelCreateOptions {
                    PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
                })) {
                    var sheet = document.AddSheet("Values");
                    sheet.Cell(1, 1).SetValue("Header");
                    sheet.Cell(2, 1).SetValue("Text");
                    sheet.Cell(3, 1).SetValue(42.5);
                    sheet.Cell(4, 1).SetValue(new DateTime(2024, 2, 29));
                    sheet.Cell(5, 1).SetValue(true);
                    sheet.Cell(6, 1).SetFormula("1+1");
                    sheet.Cell(8, 1).SetValue("Text");
                    sheet.Cell(9, 1).SetValue(-3);
                }

                package = stream.ToArray();
            }

            void AssertColumnMatches(ExcelSheet sheet, int column, int firstRow, int lastRow) {
                var snapshots = sheet.GetCellValueSnapshots(column, firstRow, lastRow);
                Assert.Equal(lastRow - firstRow + 1, snapshots.Count);
                for (int row = firstRow; row <= lastRow; row++) {
                    bool found = sheet.TryGetCellValueSnapshot(row, column, out var expected);
                    var actual = snapshots[row - firstRow];
                    Assert.Equal(found, actual != null);
                    if (!found) continue;
                    Assert.Equal(expected!.Kind, actual!.Kind);
                    Assert.Equal(expected.Text, actual.Text);
                    Assert.Equal(expected.RawValue, actual.RawValue);
                    Assert.Equal(expected.OpenXmlType, actual.OpenXmlType);
                    Assert.Equal(expected.DateTimeValue, actual.DateTimeValue);
                }
            }

            using (var loaded = ExcelDocument.Load(new MemoryStream(package))) {
                var sheet = loaded["Values"];
                AssertColumnMatches(sheet, 1, 1, 10);
                Assert.Equal(ExcelCellValueKind.DateTime, sheet.GetCellValueSnapshots(1, 4, 4)[0]!.Kind);
            }

            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Amount", typeof(double));
            table.Columns.Add("When", typeof(DateTime));
            for (int i = 0; i < 25; i++) table.Rows.Add("n" + (i % 4), i * 1.25, new DateTime(2020, 1, 1).AddDays(i));
            using var imported = ExcelDocument.Create(new MemoryStream());
            var importedSheet = imported.AddSheet("Imported");
            importedSheet.AddTable(table, new ExcelTableOptions { Name = "Imported", AutoFit = false }, row: 1, column: 1);
            for (int column = 1; column <= 3; column++) {
                AssertColumnMatches(importedSheet, column, 1, 27);
            }
        }

        [Fact]
        public void PerformanceReview_FileSnapshotKeepsAllViewsOnOneFileIdentity() {
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Snapshot.{Guid.NewGuid():N}.xlsx");
            string replacementPath = path + ".replacement";
            byte[] original = Encoding.UTF8.GetBytes("original workbook snapshot");
            byte[] replacement = Encoding.UTF8.GetBytes("replacement workbook snapshot");
            try {
                File.WriteAllBytes(path, original);
                File.WriteAllBytes(replacementPath, replacement);
                using var snapshot = SharedReadOnlyFileSnapshot.Open(path);
                using Stream firstView = snapshot.CreateView();

                File.Delete(path);
                File.Move(replacementPath, path);

                using Stream secondView = snapshot.CreateView();
                Assert.Equal(original, ReadSnapshot(firstView));
                Assert.Equal(original, ReadSnapshot(secondView));
                Assert.Equal(replacement, File.ReadAllBytes(path));
            } finally {
                File.Delete(path);
                File.Delete(replacementPath);
            }
        }

        private static byte[] ReadSnapshot(Stream stream) {
            using var output = new MemoryStream();
            stream.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] RewriteFirstWorksheetAsUtf16(byte[] packageBytes) {
            using var package = new MemoryStream();
            package.Write(packageBytes, 0, packageBytes.Length);
            package.Position = 0;

            using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
                const string worksheetPath = "xl/worksheets/sheet1.xml";
                var worksheetEntry = archive.GetEntry(worksheetPath)
                    ?? throw new InvalidOperationException("Worksheet entry was not found.");
                string worksheetXml;
                using (var reader = new StreamReader(worksheetEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
                    worksheetXml = reader.ReadToEnd();
                }

                worksheetEntry.Delete();
                var replacement = archive.CreateEntry(worksheetPath, CompressionLevel.Fastest);
                using var writer = new StreamWriter(replacement.Open(), new UnicodeEncoding(bigEndian: false, byteOrderMark: true));
                writer.Write(worksheetXml.Replace("utf-8", "utf-16").Replace("UTF-8", "utf-16"));
            }

            return package.ToArray();
        }

        private sealed class FormulaProjection {
            public FormulaProjection() {
            }

            public int Id { get; set; }
            public bool Active { get; set; }
            public string? Expression { get; set; }
        }

        private sealed class FastPackageProjection {
            internal FastPackageProjection(string name, int score) {
                Name = name;
                Score = score;
            }

            public string Name { get; }
            public int Score { get; }
        }
    }
}
