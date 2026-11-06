using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, "Net")]
        [InlineData(true, "42")]
        public void Test_CellWriteOwnership_HeaderValueRenamesSchemaAndReferences(bool numeric, string expected) {
            string path = Path.Combine(_directoryWithFiles, "HeaderValueOwnership-" + numeric + ".xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(3, 2).SetValue("Region");
                sheet.Cell(3, 3).SetValue("Amount");
                for (int row = 4; row <= 164; row++) {
                    sheet.Cell(row, 2).SetValue("EU");
                    sheet.Cell(row, 3).SetValue(10);
                }
                sheet.AddTable("B3:C164", true, "Sales", ExcelTableStyle.TableStyleMedium2);
                sheet.Cell(4, 2).SetFormula("[@Amount]*2");
                var summary = document.AddSheet("Summary");
                summary.Cell(1, 1).SetFormula("SUM(Sales[Amount])");

                if (numeric) sheet.Cell(3, 3).SetValue(42);
                else sheet.Cell(3, 3).SetValue(" Net ");

                Assert.Equal(expected, sheet.Cell(3, 3).GetValue<string>());
                Assert.Equal(new[] { "Region", expected }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
                Assert.Equal("[@" + expected + "]*2", Assert.Single(sheet.GetFormulaCells()).Formula);
                Assert.Equal("SUM(Sales[" + expected + "])", Assert.Single(summary.GetFormulaCells()).Formula);
                Assert.Equal(10d, sheet.Cell(4, 3).GetValue<double>());
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            Assert.Equal(expected, reopened["Data"].Cell(3, 3).GetValue().Value);
            Assert.Equal(expected, Assert.Single(reopened.GetTables()).Columns[1].Name);
            Assert.Equal("[@" + expected + "]*2", Assert.Single(reopened["Data"].GetFormulaCells()).Formula);
            Assert.Equal("SUM(Sales[" + expected + "])", Assert.Single(reopened["Summary"].GetFormulaCells()).Formula);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("region")]
        public void Test_CellWriteOwnership_InvalidHeaderLeavesCellSchemaAndReferencesIntact(string? value) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("Amount");
            sheet.Cell(2, 2).SetValue(10);
            sheet.AddTable("A1:B2", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("SUM(Sales[Amount])");

            Assert.Throws<ArgumentException>(() => sheet.Cell(1, 2).SetValue((object?)value));

            Assert.Equal("Amount", sheet.Cell(1, 2).GetValue<string>());
            Assert.Equal(new[] { "Region", "Amount" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
            Assert.Equal("SUM(Sales[Amount])", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Theory]
        [InlineData(ExcelExecutionMode.Sequential)]
        [InlineData(ExcelExecutionMode.Parallel)]
        public void Test_CellWriteOwnership_BulkHeadersUseFinalSchema(ExcelExecutionMode mode) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell(3, 2).SetValue("Region");
            sheet.Cell(3, 3).SetValue("Amount");
            sheet.Cell(4, 3).SetValue(10);
            sheet.AddTable("B3:C23", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(1, 1).SetFormula("SUM(Sales[Amount])");
            var writes = new List<(int Row, int Column, object Value)> {
                (3, 2, " Amount "), (3, 3, "Region")
            };
            for (int row = 4; row <= 23; row++) writes.Add((row, 3, row * 10));

            sheet.SetValues(writes, mode);

            Assert.Equal("Amount", sheet.Cell(3, 2).GetValue<string>());
            Assert.Equal("Region", sheet.Cell(3, 3).GetValue<string>());
            Assert.Equal(new[] { "Amount", "Region" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
            Assert.Equal("SUM(Sales[Region])", Assert.Single(sheet.GetFormulaCells()).Formula);
            Assert.Equal(230d, sheet.Cell(23, 3).GetValue<double>());

            Assert.Throws<ArgumentException>(() => sheet.SetValues(new[] {
                (Row: 4, Column: 3, Value: (object)999),
                (Row: 3, Column: 3, Value: (object)"amount")
            }, mode));
            Assert.Equal(40d, sheet.Cell(4, 3).GetValue<double>());
            Assert.Equal("Region", sheet.Cell(3, 3).GetValue<string>());
            Assert.Equal("SUM(Sales[Region])", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_CellWriteOwnership_HeaderRenameKeepsQueryFieldIdentity() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "SalesQuery",
                CommandText = "inert-command",
                WorksheetName = sheet.Name,
                StartCell = "B3",
                TableName = "Sales",
                ColumnNames = new[] { "Region", "Amount" }
            });
            var tablePart = Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
            var queryPart = Assert.Single(tablePart.QueryTableParts);
            uint[] fieldIds = queryPart.QueryTable!.QueryTableRefresh!.QueryTableFields!
                .Elements<QueryTableField>().Select(field => field.Id!.Value).ToArray();

            sheet.Cell(3, 3).SetValue("Net");

            var fields = queryPart.QueryTable.QueryTableRefresh.QueryTableFields!.Elements<QueryTableField>().ToArray();
            Assert.Equal(new[] { "Region", "Net" }, fields.Select(field => field.Name!.Value));
            Assert.Equal(fieldIds, fields.Select(field => field.Id!.Value));
            Assert.Equal(fieldIds, tablePart.Table!.TableColumns!.Elements<TableColumn>().Select(column => column.QueryTableFieldId!.Value));
            Assert.Equal("Net", sheet.Cell(3, 3).GetValue<string>());
            Assert.Equal("Net", Assert.Single(document.GetTables()).Columns[1].Name);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_CellWriteOwnership_DirectSchemaEditMaterializesBufferedCellsBeforeChangingHeaders(bool resize) {
            string path = Path.Combine(_directoryWithFiles, "BufferedSchemaOwnership-" + resize + ".xlsx");
            string[] expected = resize ? new[] { "Region", "Amount", "Column3" } : new[] { "Territory", "Net" };
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(3, 2).SetValue("Region");
                sheet.Cell(3, 3).SetValue("Amount");
                for (int row = 4; row <= 164; row++) {
                    sheet.Cell(row, 2).SetValue("EU");
                    sheet.Cell(row, 3).SetValue(row);
                }
                sheet.AddTable("B3:C164", true, "Sales", ExcelTableStyle.TableStyleMedium2);

                if (resize) sheet.ResizeTable("Sales", "B3:D164");
                else sheet.SetTableSchema("Sales", expected);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            var loaded = reopened["Data"];
            Assert.Equal(expected, Assert.Single(reopened.GetTables()).Columns.Select(column => column.Name));
            for (int index = 0; index < expected.Length; index++)
                Assert.Equal(expected[index], loaded.Cell(3, index + 2).GetValue<string>());
            Assert.Equal("EU", loaded.Cell(164, 2).GetValue<string>());
            Assert.Equal(164d, loaded.Cell(164, 3).GetValue<double>());
        }

        [Fact]
        public void Test_CellWriteOwnership_RichHeaderRenamesAndCaseChangesPreserveRunFormatting() {
            string path = Path.Combine(_directoryWithFiles, "RichHeaderOwnership.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(3, 2).SetRichText(new ExcelRichTextRun("Re") { Bold = true }, new ExcelRichTextRun("gion") { Italic = true });
                sheet.Cell(3, 3).SetValue("Amount");
                sheet.Cell(4, 3).SetValue(10);
                sheet.AddTable("B3:C4", true, "Sales", ExcelTableStyle.TableStyleMedium2);
                // Keep a rich sibling header after table creation as well.
                sheet.Cell(3, 2).SetRichText(new ExcelRichTextRun("Re") { Bold = true }, new ExcelRichTextRun("gion") { Italic = true });
                sheet.Cell(1, 1).SetFormula("SUM(Sales[Amount])");

                sheet.Cell(3, 3).SetRichText(new ExcelRichTextRun(" net") { Bold = true }, new ExcelRichTextRun(" Value ") { Italic = true });
                Assert.Equal("net Value", sheet.Cell(3, 3).GetValue<string>());
                Assert.Equal("net", sheet.Cell(3, 3).GetRichText()[0].Text);
                Assert.Equal(" Value", sheet.Cell(3, 3).GetRichText()[1].Text);
                sheet.Cell(3, 3).TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase, System.Globalization.CultureInfo.InvariantCulture);

                Assert.Equal("NET VALUE", Assert.Single(document.GetTables()).Columns[1].Name);
                Assert.Equal("SUM(Sales[NET VALUE])", Assert.Single(sheet.GetFormulaCells()).Formula);
                sheet.Range("B3:C3").Clear(ExcelClearOptions.Styles | ExcelClearOptions.Formulas);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            var loaded = reopened["Data"];
            var targetRuns = loaded.Cell(3, 3).GetRichText();
            Assert.Equal(new[] { "NET", " VALUE" }, targetRuns.Select(run => run.Text));
            Assert.True(targetRuns[0].Bold);
            Assert.True(targetRuns[1].Italic);
            var siblingRuns = loaded.Cell(3, 2).GetRichText();
            Assert.Equal(new[] { "Re", "gion" }, siblingRuns.Select(run => run.Text));
            Assert.True(siblingRuns[0].Bold);
            Assert.True(siblingRuns[1].Italic);
            Assert.Equal("NET VALUE", Assert.Single(reopened.GetTables()).Columns[1].Name);
            Assert.Equal("SUM(Sales[NET VALUE])", Assert.Single(loaded.GetFormulaCells()).Formula);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_CellWriteOwnership_InvalidRichHeaderChangePreservesOriginalRunsAndSchema(bool changeCase) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("i");
            sheet.Cell(1, 2).SetValue("İ");
            sheet.Cell(2, 1).SetValue(1);
            sheet.AddTable("A1:B2", true, "Labels", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(1, 1).SetRichText(new ExcelRichTextRun("i") { Bold = true });

            Assert.Throws<ArgumentException>(() => {
                if (changeCase) sheet.Cell(1, 1).TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase, System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));
                else sheet.Cell(1, 1).SetRichText(new ExcelRichTextRun("İ") { Italic = true });
            });

            Assert.Equal("i", sheet.Cell(1, 1).GetValue<string>());
            Assert.True(Assert.Single(sheet.Cell(1, 1).GetRichText()).Bold);
            Assert.Equal(new[] { "i", "İ" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
        }

        [Theory]
        [InlineData("clear-cell")]
        [InlineData("clear-range")]
        [InlineData("formula")]
        [InlineData("array")]
        [InlineData("dynamic")]
        [InlineData("error")]
        public void Test_CellWriteOwnership_NonliteralOrEmptyHeaderWritesFailBeforeMutation(string route) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("outside");
            sheet.Cell(2, 1).SetValue("Region");
            sheet.Cell(2, 2).SetValue("Amount");
            sheet.Cell(3, 1).SetValue("EU");
            sheet.Cell(3, 2).SetValue(10);
            sheet.AddTable("A2:B3", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            Action write = () => {
                switch (route) {
                    case "clear-cell": sheet.Cell(2, 1).Clear(ExcelClearOptions.Values); break;
                    case "clear-range": sheet.Range("A1:B3").Clear(); break;
                    case "formula": sheet.Cell(2, 1).SetFormula("\"Net\""); break;
                    case "array": sheet.SetArrayFormula("A1:A2", "SEQUENCE(2)"); break;
                    case "dynamic": sheet.SetDynamicArrayFormula("A2", "SEQUENCE(2)"); break;
                    case "error": sheet.Cell(2, 1).SetError("#N/A"); break;
                }
            };

            if (route.StartsWith("clear", StringComparison.Ordinal)) Assert.Throws<ArgumentException>(write);
            else Assert.Throws<InvalidOperationException>(write);

            Assert.Equal("outside", sheet.Cell(1, 1).GetValue<string>());
            Assert.Equal("Region", sheet.Cell(2, 1).GetValue<string>());
            Assert.Equal("EU", sheet.Cell(3, 1).GetValue<string>());
            Assert.Equal(10d, sheet.Cell(3, 2).GetValue<double>());
            Assert.Equal(new[] { "Region", "Amount" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
            Assert.Empty(sheet.GetFormulaCells());
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Test_CellWriteOwnership_ImportedHeaderFormulasRemainDeferredWithTheirCaches(bool array, bool omittedCoordinates) {
            string path = Path.Combine(_directoryWithFiles, "ImportedHeaderFormulaOwnership-" + array + ".xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(2, 2).SetValue(1);
                sheet.Cell(3, 2).SetValue("Region");
                sheet.Cell(3, 3).SetValue("Amount");
                sheet.Cell(4, 3).SetValue(10);
                sheet.AddTable("B3:C4", true, "Sales", ExcelTableStyle.TableStyleMedium2);
                document.Save();
            }
            using (var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true)) {
                var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell owner = worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == (array ? "B2" : "B3"));
                owner.CellFormula = array
                    ? new CellFormula("_xlfn.SEQUENCE(2)") { FormulaType = CellFormulaValues.Array, Reference = "B2:B3" }
                    : new CellFormula("\"Changed\"");
                owner.DataType = array ? CellValues.Number : CellValues.String;
                owner.CellValue = new CellValue(array ? "1" : "Region");
                if (omittedCoordinates) {
                    var row = (Row)owner.Parent!;
                    row.InsertBefore(new Cell {
                        CellReference = array ? "A2" : "A3", DataType = CellValues.String,
                        CellValue = new CellValue("left")
                    }, owner);
                    owner.CellReference = null;
                    row.RowIndex = null;
                }
                worksheet.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            var loaded = reopened["Data"];
            Assert.Equal(0, reopened.Calculate());
            Assert.Equal("Region", loaded.Cell(3, 2).GetValue<string>());
            Assert.Equal("Region", Assert.Single(reopened.GetTables()).Columns[0].Name);
            Assert.False(Assert.Single(loaded.InspectFormulas().Formulas).IsSupportedByOfficeIMO);
            Assert.Throws<InvalidOperationException>(() => loaded.ClearCachedFormulaResults());
            if (array) Assert.Throws<ArgumentException>(() => loaded.ClearArrayFormula("B2"));
            Assert.Equal("Region", loaded.Cell(3, 2).GetValue<string>());
        }

        [Fact]
        public void Test_CellWriteOwnership_MergedChildRetainsItsPhysicalValueAfterSaveAndUnmerge() {
            string path = Path.Combine(_directoryWithFiles, "MergedChildOwnership.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("anchor");
                sheet.Range("A1:B1").Merge();
                sheet.Cell(1, 2).SetValue("child");
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            var loaded = reopened["Data"];
            Assert.Equal("child", loaded.Cell(1, 2).GetValue<string>());
            loaded.Range("A1:B1").Unmerge();
            Assert.Equal("anchor", loaded.Cell(1, 1).GetValue<string>());
            Assert.Equal("child", loaded.Cell(1, 2).GetValue<string>());
        }

        [Fact]
        public void Test_CellWriteOwnership_FixedArrayChildCachePersistsUntilExplicitCalculation() {
            string path = Path.Combine(_directoryWithFiles, "FixedArrayChildOwnership.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.SetArrayFormula("G1:H2", "SEQUENCE(2,2)");
                sheet.Cell(2, 8).SetValue(99);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            var loaded = reopened["Data"];
            Assert.Equal(99d, loaded.Cell(2, 8).GetValue<double>());
            Assert.Equal("G1:H2", Assert.Single(loaded.InspectFormulas().Formulas).Array!.Range);
            Assert.Equal(1, reopened.Calculate());
            Assert.Equal(4d, loaded.Cell(2, 8).GetValue<double>());
        }
    }
}
