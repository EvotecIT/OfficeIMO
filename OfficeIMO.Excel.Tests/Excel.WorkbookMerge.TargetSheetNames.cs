using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelWorkbookMerge_TargetSheetNamesMapEachSourceNameDirectly() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TargetNames.Source.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TargetNames.Target.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet a = source.AddSheet("A");
                ExcelSheet xa = source.AddSheet("X - A");
                a.Cell(1, 1).SetValue(1);
                a.Cell(1, 2).SetFormula("'X - A'!A1");
                a.Cell(2, 2).SetFormula("SUM(A!A1,'X - A'!A1)+SUM(A:'X - A'!A1)+Lookup");
                a.Cell(1, 3).SetFormula("'X - A'!A1");
                a.Cell(2, 3).SetFormula("'X - A'!A2");
                a.SetInternalLink(3, 1, "'X - A'!A1", "Go");

                xa.Cell(1, 1).SetValue(2);
                xa.Cell(1, 2).SetFormula("A!A1");
                xa.Range("C1").Validation().CustomFormula("COUNTIF(A!$A$1:$A$1,\">0\")>0");
                source.Save();
            }

            AddWorkbookDefinedName(sourcePath, "Lookup", "'X - A'!$A$1");
            MakeSharedFormula(sourcePath, "A", "C1", "C2");

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    TargetSheetNames = new Dictionary<string, string> {
                        ["A"] = "X - A",
                        ["X - A"] = "X - X - A"
                    },
                    SheetNameValidationMode = ExcelSheetNameValidationMode.Strict,
                    CopyMode = ExcelWorksheetCopyMode.Package
                });

                Assert.Equal(new[] { "A", "X - A" }, result.SourceSheets);
                Assert.Equal(new[] { "X - A", "X - X - A" }, result.TargetSheets);
                target.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
                Assert.Equal(new[] { "X - A", "X - X - A" }, spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Select(_ => _.Name!.Value));

                WorksheetPart fromA = GetWorksheetPartByNameForOperations(spreadsheet, "X - A");
                WorksheetPart fromXa = GetWorksheetPartByNameForOperations(spreadsheet, "X - X - A");
                Assert.Equal("1", GetMergedCell(fromA, "A1").CellValue?.Text);
                Assert.Equal("'X - X - A'!A1", GetMergedCell(fromA, "B1").CellFormula?.Text);
                Assert.Equal("SUM('X - A'!A1,'X - X - A'!A1)+SUM('X - A:X - X - A'!A1)+Lookup", GetMergedCell(fromA, "B2").CellFormula?.Text);
                Assert.Equal("'X - X - A'!A1", Assert.Single(fromA.Worksheet.Descendants<Hyperlink>()).Location?.Value);

                CellFormula sharedMaster = GetMergedCell(fromA, "C1").CellFormula!;
                CellFormula sharedChild = GetMergedCell(fromA, "C2").CellFormula!;
                Assert.Equal(CellFormulaValues.Shared, sharedMaster.FormulaType?.Value);
                Assert.Equal("'X - X - A'!A1", sharedMaster.Text);
                Assert.Equal(sharedMaster.SharedIndex?.Value, sharedChild.SharedIndex?.Value);
                Assert.True(string.IsNullOrEmpty(sharedChild.Text));

                Assert.Equal("2", GetMergedCell(fromXa, "A1").CellValue?.Text);
                Assert.Equal("'X - A'!A1", GetMergedCell(fromXa, "B1").CellFormula?.Text);
                Assert.Equal("COUNTIF('X - A'!$A$1:$A$1,\">0\")>0", Assert.Single(fromXa.Worksheet.Descendants<Formula1>()).Text);

                DefinedName lookup = Assert.Single(spreadsheet.WorkbookPart.Workbook.DefinedNames!.Elements<DefinedName>(), _ => _.Name == "Lookup");
                Assert.Equal("'X - X - A'!$A$1", lookup.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_TargetSheetNamesQuoteApostrophesAndLeaveStringLiterals() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TargetNamesQuoting.Source.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TargetNamesQuoting.Target.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddSheet("Data's Sheet").Cell(1, 1).SetValue(5);
                source.AddSheet("Plain").Cell(1, 1).SetValue(6);
                ExcelSheet summary = source.AddSheet("Summary");
                summary.Cell(1, 1).SetFormula("'Data''s Sheet'!A1&\"'Data''s Sheet'!A1\"");
                summary.Cell(2, 1).SetFormula("Plain!A1&\"Plain!A1\"&\"\"\"Plain!A1\"\"\"");
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    TargetSheetNames = new Dictionary<string, string> {
                        ["Data's Sheet"] = "Bob's Data",
                        ["Plain"] = "Brief's Plain",
                        ["Summary"] = "Brief Summary"
                    },
                    SheetNameValidationMode = ExcelSheetNameValidationMode.Strict,
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                target.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
                WorksheetPart summaryPart = GetWorksheetPartByNameForOperations(spreadsheet, "Brief Summary");
                Assert.Equal("'Bob''s Data'!A1&\"'Data''s Sheet'!A1\"", GetMergedCell(summaryPart, "A1").CellFormula?.Text);
                Assert.Equal("'Brief''s Plain'!A1&\"Plain!A1\"&\"\"\"Plain!A1\"\"\"", GetMergedCell(summaryPart, "A2").CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_TargetSheetNamesRewriteTableCalculatedColumns() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TargetNamesTable.Source.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TargetNamesTable.Target.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddSheet("Rates").Cell(1, 1).SetValue(0.2);
                ExcelSheet orders = source.AddSheet("Orders");
                orders.Cell(1, 1).SetValue("Amount");
                orders.Cell(1, 2).SetValue("Tax");
                orders.Cell(2, 1).SetValue(100);
                orders.Cell(2, 2).SetFormula("OrderLines[[#This Row],[Amount]]*Rates!$A$1");
                orders.AddTable("A1:B2", hasHeader: true, name: "OrderLines", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium9);
                source.Save();
            }

            AddTableCalculatedColumnFormula(sourcePath, "Orders", "Tax", "OrderLines[[#This Row],[Amount]]*Rates!$A$1");

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    TargetSheetNames = new Dictionary<string, string> {
                        ["Rates"] = "Brief - Rates",
                        ["Orders"] = "Brief - Orders"
                    },
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                target.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
                WorksheetPart ordersPart = GetWorksheetPartByNameForOperations(spreadsheet, "Brief - Orders");
                Table table = Assert.Single(ordersPart.TableDefinitionParts).Table!;
                CalculatedColumnFormula calculated = Assert.Single(table.Descendants<CalculatedColumnFormula>());
                Assert.Equal("OrderLines[[#This Row],[Amount]]*'Brief - Rates'!$A$1", calculated.Text);
                Assert.Equal("OrderLines[[#This Row],[Amount]]*'Brief - Rates'!$A$1", GetMergedCell(ordersPart, "B2").CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_UnmappedSheetsKeepPrefixedNames() {
            using var sourceStream = new MemoryStream();
            using (var source = ExcelDocument.Create(sourceStream)) {
                source.AddSheet("North").Cell(1, 1).SetValue("North value");
                source.AddSheet("South").Cell(1, 1).SetValue("South value");
                source.Save(sourceStream);
            }

            sourceStream.Position = 0;
            using var loaded = ExcelDocument.Load(sourceStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            using var mappedStream = new MemoryStream();
            using (var target = ExcelDocument.Create(mappedStream)) {
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(loaded, new ExcelWorkbookMergeOptions {
                    SheetNamePrefix = "Imported ",
                    TargetSheetNames = new Dictionary<string, string> { ["north"] = "Region N" }
                });

                Assert.Equal(new[] { "Region N", "Imported South" }, result.TargetSheets);
                Assert.True(target["Region N"].TryGetCellText(1, 1, out var northValue));
                Assert.Equal("North value", northValue);
            }

            using var emptyMapStream = new MemoryStream();
            using (var target = ExcelDocument.Create(emptyMapStream)) {
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(loaded, new ExcelWorkbookMergeOptions {
                    SheetNamePrefix = "Imported ",
                    TargetSheetNames = new Dictionary<string, string>()
                });

                Assert.Equal(new[] { "Imported North", "Imported South" }, result.TargetSheets);
            }
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_TargetSheetNamesCloneIndependently() {
            var targetSheetNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["North"] = "Region N" };
            var options = new ExcelWorkbookMergeOptions { TargetSheetNames = targetSheetNames };

            ExcelWorkbookMergeOptions clone = options.Clone();
            targetSheetNames["North"] = "Changed";

            Assert.Equal("Region N", clone.TargetSheetNames!["North"]);
            Assert.Equal("Region N", clone.TargetSheetNames["north"]);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_StrictTargetSheetNamesFailBeforeImporting() {
            using var sourceStream = new MemoryStream();
            using (var source = ExcelDocument.Create(sourceStream)) {
                source.AddSheet("A").Cell(1, 1).SetValue(1);
                source.AddSheet("B").Cell(1, 1).SetValue(2);
                source.Save(sourceStream);
            }

            sourceStream.Position = 0;
            using var loaded = ExcelDocument.Load(sourceStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            using var targetStream = new MemoryStream();
            using var target = ExcelDocument.Create(targetStream);
            target.AddSheet("Summary");

            ArgumentException existing = Assert.Throws<ArgumentException>(() => target.MergeWorkbookFrom(loaded, StrictPackageMerge(new Dictionary<string, string> {
                ["A"] = "Fine",
                ["B"] = "summary"
            })));
            Assert.Contains("already exists", existing.Message, StringComparison.Ordinal);

            ArgumentException duplicate = Assert.Throws<ArgumentException>(() => target.MergeWorkbookFrom(loaded, StrictPackageMerge(new Dictionary<string, string> {
                ["A"] = "Same",
                ["B"] = "same"
            })));
            Assert.Contains("more than one imported worksheet", duplicate.Message, StringComparison.Ordinal);

            ArgumentException tooLong = Assert.Throws<ArgumentException>(() => target.MergeWorkbookFrom(loaded, StrictPackageMerge(new Dictionary<string, string> {
                ["B"] = new string('B', 32)
            })));
            Assert.Contains("31 characters", tooLong.Message, StringComparison.Ordinal);

            ArgumentException unknown = Assert.Throws<ArgumentException>(() => target.MergeWorkbookFrom(loaded, StrictPackageMerge(new Dictionary<string, string> {
                ["Missing"] = "Anything"
            })));
            Assert.Contains("'Missing'", unknown.Message, StringComparison.Ordinal);

            Assert.Equal(new[] { "Summary" }, target.Sheets.Select(_ => _.Name));
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_SanitizedTargetSheetNamesRewriteToFinalNames() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TargetNamesSanitize.Source.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddSheet("A").Cell(1, 1).SetValue(1);
                ExcelSheet b = source.AddSheet("B");
                b.Cell(1, 1).SetFormula("A!A1");
                ExcelSheet c = source.AddSheet("C");
                c.Cell(1, 1).SetFormula("B!A1");
                source.Save();
            }

            string longName = new string('L', 40);
            using var targetStream = new MemoryStream();
            using var target = ExcelDocument.Create(targetStream);
            using var loaded = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            target.AddSheet("Summary");

            ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(loaded, new ExcelWorkbookMergeOptions {
                TargetSheetNames = new Dictionary<string, string> {
                    ["A"] = "Summary",
                    ["B"] = longName
                },
                CopyMode = ExcelWorksheetCopyMode.Package
            });

            string truncatedName = longName.Substring(0, 31);
            Assert.Equal(new[] { "Summary (2)", truncatedName, "C" }, result.TargetSheets);
            Assert.Equal("'Summary (2)'!A1", target[truncatedName].WorksheetPart.Worksheet!.Descendants<CellFormula>().Single().Text);
            Assert.Equal($"'{truncatedName}'!A1", target["C"].WorksheetPart.Worksheet!.Descendants<CellFormula>().Single().Text);

            File.Delete(sourcePath);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_KeepsOnlyActiveWorksheetTabSelected() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TabSelection.Source.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.TabSelection.Target.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddSheet("One").Cell(1, 1).SetValue(1);
                source.AddSheet("Two").Cell(1, 1).SetValue(2);
                source.Save();
            }

            SelectTab(sourcePath, "One");
            SelectTab(sourcePath, "Two");

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                target.AddSheet("Summary");
                target.SetActiveWorksheet("Summary");
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions { CopyMode = ExcelWorksheetCopyMode.Package });
                target.Save();
            }

            Assert.Equal(new[] { "Summary" }, GetTabSelectedSheetNames(targetPath));

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions { CopyMode = ExcelWorksheetCopyMode.Package });
                target.Save();
            }

            Assert.Equal(new[] { "One" }, GetTabSelectedSheetNames(targetPath));
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        private static ExcelWorkbookMergeOptions StrictPackageMerge(IReadOnlyDictionary<string, string> targetSheetNames) {
            return new ExcelWorkbookMergeOptions {
                TargetSheetNames = targetSheetNames,
                SheetNameValidationMode = ExcelSheetNameValidationMode.Strict,
                CopyMode = ExcelWorksheetCopyMode.Package
            };
        }

        private static Cell GetMergedCell(WorksheetPart worksheetPart, string reference) {
            return worksheetPart.Worksheet.Descendants<Cell>().Single(_ => _.CellReference?.Value == reference);
        }

        private static void MakeSharedFormula(string path, string sheetName, string masterReference, string childReference) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            CellFormula master = GetMergedCell(worksheetPart, masterReference).CellFormula!;
            master.FormulaType = CellFormulaValues.Shared;
            master.Reference = masterReference + ":" + childReference;
            master.SharedIndex = 0U;
            Cell child = GetMergedCell(worksheetPart, childReference);
            child.CellFormula = new CellFormula {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 0U
            };
            worksheetPart.Worksheet.Save();
        }

        private static void SelectTab(string path, string sheetName) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            Worksheet worksheet = GetWorksheetPartByNameForOperations(spreadsheet, sheetName).Worksheet;
            SheetViews? sheetViews = worksheet.GetFirstChild<SheetViews>();
            if (sheetViews == null) {
                sheetViews = new SheetViews();
                worksheet.InsertBefore(sheetViews, worksheet.GetFirstChild<SheetData>());
            }

            SheetView? sheetView = sheetViews.GetFirstChild<SheetView>();
            if (sheetView == null) {
                sheetView = new SheetView { WorkbookViewId = 0U };
                sheetViews.Append(sheetView);
            }

            sheetView.TabSelected = true;
            worksheet.Save();
        }

        private static string[] GetTabSelectedSheetNames(string path) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, false);
            return spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
                .Where(_ => GetWorksheetPartByNameForOperations(spreadsheet, _.Name!.Value!).Worksheet
                    .Descendants<SheetView>()
                    .Any(_ => _.TabSelected?.Value == true))
                .Select(_ => _.Name!.Value!)
                .ToArray();
        }
    }
}
