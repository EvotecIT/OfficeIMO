using System.IO;
using System.Linq;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ReorderWorkSheet_PersistsWorkbookOrder() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetReorder.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Alpha");
                document.AddWorkSheet("Beta");
                document.AddWorkSheet("Gamma");

                document.ReorderWorkSheet("Gamma", 0);

                Assert.Equal(new[] { "Gamma", "Alpha", "Beta" }, document.GetSheetNames());
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                string[] names = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
                    .Select(sheet => sheet.Name?.Value ?? string.Empty)
                    .ToArray();
                Assert.Equal(new[] { "Gamma", "Alpha", "Beta" }, names);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CopyWorkSheetWithinWorkbook_CopiesValuesAndSanitizesName() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyWithin.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorkSheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "Score");
                source.CellValue(2, 1, "Ada");
                source.CellValue(2, 2, 10);

                ExcelSheet copy = document.CopyWorkSheet(source, "Copy:Source");

                Assert.Equal("Copy_Source", copy.Name);
                Assert.Equal("A1:B2", copy.GetUsedRangeA1());
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Equal(new[] { "Source", "Copy_Source" }, document.GetSheetNames());
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Copy_Source").ReadRange("A1:B2");
                Assert.Equal("Name", values[0, 0]);
                Assert.Equal("Score", values[0, 1]);
                Assert.Equal("Ada", values[1, 0]);
                Assert.Equal(10d, values[1, 1]);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CopyWorkSheetWithinWorkbook_PreservesTablesWithUniqueNames() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyWithinTables.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorkSheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "NA");
                source.CellValue(2, 2, 100);
                source.CellValue(3, 1, "EMEA");
                source.CellValue(3, 2, 200);
                source.AddTable("A1:B3", hasHeader: true, name: "SalesTable", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellFormula(4, 2, "SUM(SalesTable[Revenue])");

                ExcelSheet copy = document.CopyWorkSheet(source, "Copy");

                Assert.Equal("A1:B3", copy.GetTableRange("SalesTable2"));
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.DocumentValidationErrors);
                Assert.Equal("A1:B3", document.GetSheet("Copy").GetTableRange("SalesTable2"));
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Copy");
                TableDefinitionPart copiedTablePart = Assert.Single(copiedPart.TableDefinitionParts);
                Table copiedTable = copiedTablePart.Table;
                Assert.Equal("SalesTable2", copiedTable.Name?.Value);
                Assert.Equal("SalesTable2", copiedTable.DisplayName?.Value);
                Assert.Equal("A1:B3", copiedTable.Reference?.Value);
                Assert.Equal("TableStyleMedium9", copiedTable.TableStyleInfo?.Name?.Value);

                TableParts tableParts = Assert.Single(copiedPart.Worksheet.Elements<TableParts>());
                TablePart tablePart = Assert.Single(tableParts.Elements<TablePart>());
                Assert.NotNull(copiedPart.GetPartById(tablePart.Id!.Value!));

                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "B4");
                Assert.Equal("SUM(SalesTable2[Revenue])", formulaCell.CellFormula?.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CopyWorkSheetWithinWorkbook_InsertsTablePartsBeforeExtensionList() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyTablePartsOrder.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorkSheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Ada");
                source.AddTable("A1:A2", hasHeader: true, name: "People", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(new WorksheetExtension { Uri = "{00000000-0000-0000-0000-000000000001}" }));

                document.CopyWorkSheet(source, "Copy");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Copy");
                var children = copiedPart.Worksheet.ChildElements.ToList();
                int tablePartsIndex = children.FindIndex(element => element is TableParts);
                int extensionListIndex = children.FindIndex(element => element is WorksheetExtensionList);

                Assert.True(tablePartsIndex >= 0);
                Assert.True(extensionListIndex >= 0);
                Assert.True(tablePartsIndex < extensionListIndex);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CopyWorkSheetWithinWorkbook_RewritesStructuredReferencesAtomicallyOutsideStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyStructuredReferenceRewrite.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorkSheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "NA");
                source.CellValue(2, 2, 100);
                source.CellValue(1, 4, "Region");
                source.CellValue(1, 5, "Revenue");
                source.CellValue(2, 4, "EMEA");
                source.CellValue(2, 5, 200);
                source.AddTable("A1:B2", hasHeader: true, name: "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.AddTable("D1:E2", hasHeader: true, name: "Sales2", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellFormula(4, 1, "SUM(Sales[Revenue])+SUM(Sales2[Revenue])+\"Sales[Revenue]\"");

                document.CopyWorkSheet(source, "Copy");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Copy");
                var tableNamesByRange = copiedPart.TableDefinitionParts
                    .Select(part => part.Table)
                    .Where(table => table?.Reference?.Value != null && table.Name?.Value != null)
                    .ToDictionary(table => table!.Reference!.Value!, table => table!.Name!.Value!);

                string firstCopiedTable = tableNamesByRange["A1:B2"];
                string secondCopiedTable = tableNamesByRange["D1:E2"];
                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "A4");

                Assert.Equal($"SUM({firstCopiedTable}[Revenue])+SUM({secondCopiedTable}[Revenue])+\"Sales[Revenue]\"", formulaCell.CellFormula?.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CopyWorkSheetFrom_CopiesValuesBetweenWorkbooks() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopySource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorkSheet("Source");
                source.CellValue(2, 2, "Region");
                source.CellValue(2, 3, "Revenue");
                source.CellValue(3, 2, "NA");
                source.CellValue(3, 3, 125.5m);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, readOnly: true))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet copied = targetDocument.CopyWorkSheetFrom(sourceDocument, "Source", "Imported");

                Assert.Equal("Imported", copied.Name);
                Assert.Equal("B2:C3", copied.GetUsedRangeA1());
                targetDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, readOnly: true)) {
                using var reader = targetDocument.CreateReader();
                object?[,] values = reader.GetSheet("Imported").ReadRange("B2:C3");
                Assert.Equal("Region", values[0, 0]);
                Assert.Equal("Revenue", values[0, 1]);
                Assert.Equal("NA", values[1, 0]);
                Assert.Equal(125.5d, values[1, 1]);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorkSheetFrom_PreservesTablesBetweenWorkbooks() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyTableSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyTableTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorkSheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "NA");
                source.CellValue(2, 2, 100);
                source.CellValue(3, 1, "EMEA");
                source.CellValue(3, 2, 200);
                source.AddTable("A1:B3", hasHeader: true, name: "SourceSales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, readOnly: true))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet copied = targetDocument.CopyWorkSheetFrom(sourceDocument, "Source", "Imported");

                Assert.Equal("A1:B3", copied.GetTableRange("SourceSales"));
                targetDocument.Save();
            }

            using (var document = ExcelDocument.Load(targetPath, readOnly: true)) {
                Assert.Empty(document.DocumentValidationErrors);
                Assert.Equal("A1:B3", document.GetSheet("Imported").GetTableRange("SourceSales"));
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                TableDefinitionPart copiedTablePart = Assert.Single(copiedPart.TableDefinitionParts);
                Table copiedTable = copiedTablePart.Table;
                Assert.Equal("SourceSales", copiedTable.Name?.Value);
                Assert.Equal("A1:B3", copiedTable.Reference?.Value);
                Assert.Equal("TableStyleMedium9", copiedTable.TableStyleInfo?.Name?.Value);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorkSheetFrom_PackageModePreservesHeaderOnlyStyles() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageHeaderSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageHeaderTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorkSheet("Headers");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellBold(1, 1, true);
                source.CellBold(1, 2, true);
                source.CellBackground(1, 1, "#D9EAD3");
                source.CellBackground(1, 2, "#D9EAD3");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, readOnly: true))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorkSheet("Summary").CellValue(1, 1, "Summary");
                ExcelSheet copied = targetDocument.CopyWorkSheetFrom(sourceDocument, "Headers", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });

                Assert.Equal("Imported", copied.Name);
                Assert.Equal("A1:B1", copied.GetUsedRangeA1());
                targetDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, readOnly: true)) {
                Assert.True(targetDocument["Imported"].TryGetCellText(1, 1, out var header));
                Assert.Equal("Region", header);
                Assert.Equal("A1:B1", targetDocument["Imported"].GetUsedRangeA1());
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell headerCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Assert.Equal(CellValues.InlineString, headerCell.DataType?.Value);
                Assert.NotNull(headerCell.StyleIndex);
                Assert.NotEqual(0U, headerCell.StyleIndex!.Value);

                Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                CellFormat headerFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)headerCell.StyleIndex!.Value);
                Assert.True(headerFormat.ApplyFont?.Value ?? false);
                Assert.True(headerFormat.ApplyFill?.Value ?? false);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorkSheetFrom_PackageModeDataTableReadsInlineStrings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageInlineSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageInlineTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorkSheet("External");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Imported");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, readOnly: true))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorkSheetFrom(sourceDocument, "External", "ExternalCopy", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (var reader = ExcelDocumentReader.Open(targetPath)) {
                var table = reader.GetSheet("ExternalCopy").ReadRangeAsDataTable("A1:A2");
                Assert.Single(table.Columns);
                Assert.Equal("Name", table.Columns[0].ColumnName);
                DataRow row = Assert.Single(table.Rows.Cast<DataRow>());
                Assert.Equal("Imported", row["Name"]);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorkSheetFrom_ValuesModeKeepsReaderWriterFallback() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyValuesModeSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyValuesModeTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorkSheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Ada");
                source.CellBold(1, 1, true);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, readOnly: true))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet copied = targetDocument.CopyWorkSheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Values
                });

                Assert.Equal("Imported", copied.Name);
                Assert.Equal("A1:A2", copied.GetUsedRangeA1());
                targetDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, readOnly: true)) {
                Assert.True(targetDocument["Imported"].TryGetCellText(2, 1, out var value));
                Assert.Equal("Ada", value);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell headerCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Assert.True(headerCell.StyleIndex == null || headerCell.StyleIndex.Value == 0U);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CompareRanges_ReturnsCellDifferences() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCompare.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet left = document.AddWorkSheet("Left");
                left.CellValue(1, 1, "Name");
                left.CellValue(1, 2, "Score");
                left.CellValue(2, 1, "Ada");
                left.CellValue(2, 2, 10);

                ExcelSheet right = document.AddWorkSheet("Right");
                right.CellValue(1, 1, "Name");
                right.CellValue(1, 2, "Score");
                right.CellValue(2, 1, "Ada");
                right.CellValue(2, 2, 11);

                var differences = document.CompareRanges(left, "A1:B2", right, "A1:B2");

                ExcelRangeDifference difference = Assert.Single(differences);
                Assert.Equal(ExcelRangeDifferenceKind.ValueMismatch, difference.Kind);
                Assert.Equal("B2", difference.LeftCell);
                Assert.Equal("B2", difference.RightCell);
                Assert.Equal(10d, difference.LeftValue);
                Assert.Equal(11d, difference.RightValue);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_MergeWorkSheets_AppendsRowsAndSkipsSourceHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeAppend.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorkSheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");
                target.CellValue(2, 1, "NA");
                target.CellValue(2, 2, 100);

                ExcelSheet source = document.AddWorkSheet("More");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "EMEA");
                source.CellValue(2, 2, 200);
                source.CellValue(3, 1, "APAC");
                source.CellValue(3, 2, 150);

                ExcelWorksheetMergeResult result = document.MergeWorkSheets(target, source);

                Assert.Equal("Combined", result.TargetSheetName);
                Assert.Equal("More", result.SourceSheetName);
                Assert.Equal("A1:B3", result.SourceRange);
                Assert.Equal("A3:B4", result.TargetRange);
                Assert.Equal(2, result.RowsCopied);
                Assert.Equal(2, result.ColumnsCopied);
                Assert.True(result.HeaderSkipped);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Combined").ReadRange("A1:B4");
                Assert.Equal("Region", values[0, 0]);
                Assert.Equal("Revenue", values[0, 1]);
                Assert.Equal("NA", values[1, 0]);
                Assert.Equal(100d, values[1, 1]);
                Assert.Equal("EMEA", values[2, 0]);
                Assert.Equal(200d, values[2, 1]);
                Assert.Equal("APAC", values[3, 0]);
                Assert.Equal(150d, values[3, 1]);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_MergeWorkSheets_CanMatchColumnsByHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeHeaders.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorkSheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");
                target.CellValue(2, 1, "NA");
                target.CellValue(2, 2, 100);

                ExcelSheet source = document.AddWorkSheet("More");
                source.CellValue(1, 1, "Revenue");
                source.CellValue(1, 2, "Region");
                source.CellValue(2, 1, 200);
                source.CellValue(2, 2, "EMEA");
                source.CellValue(3, 1, 150);
                source.CellValue(3, 2, "APAC");

                ExcelWorksheetMergeResult result = document.MergeWorkSheets(target, source, new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true
                });

                Assert.Equal("A3:B4", result.TargetRange);
                Assert.Equal(2, result.RowsCopied);
                Assert.Equal(2, result.ColumnsCopied);
                Assert.True(result.HeaderSkipped);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Combined").ReadRange("A1:B4");
                Assert.Equal("Region", values[0, 0]);
                Assert.Equal("Revenue", values[0, 1]);
                Assert.Equal("NA", values[1, 0]);
                Assert.Equal(100d, values[1, 1]);
                Assert.Equal("EMEA", values[2, 0]);
                Assert.Equal(200d, values[2, 1]);
                Assert.Equal("APAC", values[3, 0]);
                Assert.Equal(150d, values[3, 1]);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_MergeWorkSheets_CanMatchColumnsUsingExplicitTargetHeaderRow() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeExplicitHeaderRow.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorkSheet("Combined");
                target.CellValue(1, 1, "Quarterly report");
                target.CellValue(3, 2, "Region");
                target.CellValue(3, 3, "Revenue");
                target.CellValue(4, 2, "NA");
                target.CellValue(4, 3, 100);

                ExcelSheet source = document.AddWorkSheet("More");
                source.CellValue(1, 1, "Revenue");
                source.CellValue(1, 2, "Region");
                source.CellValue(2, 1, 200);
                source.CellValue(2, 2, "EMEA");

                ExcelWorksheetMergeResult result = document.MergeWorkSheets(target, source, new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true,
                    TargetHeaderRow = 3,
                    TargetStartRow = 5,
                    TargetStartColumn = 2
                });

                Assert.Equal("B5:C5", result.TargetRange);
                Assert.Equal(1, result.RowsCopied);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Combined").ReadRange("B3:C5");
                Assert.Equal("Region", values[0, 0]);
                Assert.Equal("Revenue", values[0, 1]);
                Assert.Equal("NA", values[1, 0]);
                Assert.Equal(100d, values[1, 1]);
                Assert.Equal("EMEA", values[2, 0]);
                Assert.Equal(200d, values[2, 1]);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_MergeWorkSheets_HeaderMatchThrowsWhenSourceColumnIsMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeHeadersMissing.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorkSheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");
                target.CellValue(2, 1, "NA");
                target.CellValue(2, 2, 100);

                ExcelSheet source = document.AddWorkSheet("More");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Amount");
                source.CellValue(2, 1, "EMEA");
                source.CellValue(2, 2, 200);

                var exception = Assert.Throws<ArgumentException>(() => document.MergeWorkSheets(target, source, new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true
                }));
                Assert.Contains("Revenue", exception.Message);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_MergeWorkSheets_AllowsSourceFromAnotherWorkbook() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetMergeExternalSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetMergeExternalTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorkSheet("More");
                source.CellValue(1, 1, "Revenue");
                source.CellValue(1, 2, "Region");
                source.CellValue(2, 1, 200);
                source.CellValue(2, 2, "EMEA");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, readOnly: true))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet target = targetDocument.AddWorkSheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");

                ExcelWorksheetMergeResult result = targetDocument.MergeWorkSheets(target, sourceDocument.GetSheet("More"), new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true
                });

                Assert.Equal("A2:B2", result.TargetRange);
                Assert.Equal(1, result.RowsCopied);
                targetDocument.Save();
            }

            using (var document = ExcelDocument.Load(targetPath, readOnly: true)) {
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Combined").ReadRange("A1:B2");
                Assert.Equal("Region", values[0, 0]);
                Assert.Equal("Revenue", values[0, 1]);
                Assert.Equal("EMEA", values[1, 0]);
                Assert.Equal(200d, values[1, 1]);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_MergeWorkSheets_ThrowsWhenTargetBelongsToAnotherWorkbook() {
            string leftPath = Path.Combine(_directoryWithFiles, "WorksheetMergeWrongTargetLeft.xlsx");
            string rightPath = Path.Combine(_directoryWithFiles, "WorksheetMergeWrongTargetRight.xlsx");

            using (var leftDocument = ExcelDocument.Create(leftPath))
            using (var rightDocument = ExcelDocument.Create(rightPath)) {
                ExcelSheet wrongTarget = rightDocument.AddWorkSheet("WrongTarget");
                wrongTarget.CellValue(1, 1, "Region");

                ExcelSheet source = leftDocument.AddWorkSheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(2, 1, "EMEA");

                var exception = Assert.Throws<ArgumentException>(() => leftDocument.MergeWorkSheets(wrongTarget, source));
                Assert.Contains("Target worksheet", exception.Message);
            }

            File.Delete(leftPath);
            File.Delete(rightPath);
        }

        [Fact]
        public void Test_JoinWorksheets_CopiesRequestedRangeToExplicitPosition() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetJoinRange.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorkSheet("Target");
                ExcelSheet source = document.AddWorkSheet("Source");
                source.CellValue(1, 1, "ignore");
                source.CellValue(2, 2, "Name");
                source.CellValue(2, 3, "Score");
                source.CellValue(3, 2, "Ada");
                source.CellValue(3, 3, 10);

                ExcelWorksheetMergeResult result = document.JoinWorksheets(target, source, new ExcelWorksheetMergeOptions {
                    SourceRange = "B2:C3",
                    IncludeSourceHeader = true,
                    TargetStartRow = 5,
                    TargetStartColumn = 4
                });

                Assert.Equal("D5:E6", result.TargetRange);
                Assert.Equal(2, result.RowsCopied);
                Assert.Equal(2, result.ColumnsCopied);
                Assert.False(result.HeaderSkipped);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Target").ReadRange("D5:E6");
                Assert.Equal("Name", values[0, 0]);
                Assert.Equal("Score", values[0, 1]);
                Assert.Equal("Ada", values[1, 0]);
                Assert.Equal(10d, values[1, 1]);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_JoinWorksheets_ThrowsWhenExplicitTargetWouldOverwriteCell() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetJoinOverwriteBlocked.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorkSheet("Target");
                target.CellValue(5, 4, "Existing");

                ExcelSheet source = document.AddWorkSheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "Score");
                source.CellValue(2, 1, "Ada");
                source.CellValue(2, 2, 10);

                var exception = Assert.Throws<InvalidOperationException>(() => document.JoinWorksheets(target, source, new ExcelWorksheetMergeOptions {
                    SourceRange = "A1:B2",
                    IncludeSourceHeader = true,
                    TargetStartRow = 5,
                    TargetStartColumn = 4
                }));

                Assert.Contains("D5", exception.Message);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_JoinWorksheets_CanOverwriteExplicitTargetWhenEnabled() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetJoinOverwriteEnabled.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorkSheet("Target");
                target.CellValue(5, 4, "Existing");

                ExcelSheet source = document.AddWorkSheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "Score");
                source.CellValue(2, 1, "Ada");
                source.CellValue(2, 2, 10);

                ExcelWorksheetMergeResult result = document.JoinWorksheets(target, source, new ExcelWorksheetMergeOptions {
                    SourceRange = "A1:B2",
                    IncludeSourceHeader = true,
                    TargetStartRow = 5,
                    TargetStartColumn = 4,
                    OverwriteExistingCells = true
                });

                Assert.Equal("D5:E6", result.TargetRange);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Target").ReadRange("D5:E6");
                Assert.Equal("Name", values[0, 0]);
                Assert.Equal("Score", values[0, 1]);
                Assert.Equal("Ada", values[1, 0]);
                Assert.Equal(10d, values[1, 1]);
            }

            File.Delete(filePath);
        }

        private static WorksheetPart GetWorksheetPartByNameForOperations(SpreadsheetDocument document, string sheetName) {
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>()
                .Single(candidate => candidate.Name?.Value == sheetName);
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        }
    }
}
