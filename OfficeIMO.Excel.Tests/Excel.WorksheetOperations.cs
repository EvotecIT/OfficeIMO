using System.IO;
using System.Linq;
using System.Data;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ReorderWorksheet_PersistsWorkbookOrder() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetReorder.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Alpha");
                document.AddWorksheet("Beta");
                document.AddWorksheet("Gamma");

                document.ReorderWorksheet("Gamma", 0);

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
        public void Test_CopyWorksheetWithinWorkbook_CopiesValuesAndSanitizesName() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyWithin.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "Score");
                source.CellValue(2, 1, "Ada");
                source.CellValue(2, 2, 10);

                ExcelSheet copy = document.CopyWorksheet(source, "Copy:Source");

                Assert.Equal("Copy_Source", copy.Name);
                Assert.Equal("A1:B2", copy.GetUsedRangeA1());
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_CopyWorksheetWithinWorkbook_PreservesTablesWithUniqueNames() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyWithinTables.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "NA");
                source.CellValue(2, 2, 100);
                source.CellValue(3, 1, "EMEA");
                source.CellValue(3, 2, 200);
                source.AddTable("A1:B3", hasHeader: true, name: "SalesTable", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellFormula(4, 2, "SUM(SalesTable[Revenue])");

                ExcelSheet copy = document.CopyWorksheet(source, "Copy");

                Assert.Equal("A1:B3", copy.GetTableRange("SalesTable2"));
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_CopyWorksheetWithinWorkbook_InsertsTablePartsBeforeExtensionList() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyTablePartsOrder.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Ada");
                source.AddTable("A1:A2", hasHeader: true, name: "People", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(new WorksheetExtension { Uri = "{00000000-0000-0000-0000-000000000001}" }));

                document.CopyWorksheet(source, "Copy");
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
        public void Test_CopyWorksheetWithinWorkbook_RewritesStructuredReferencesAtomicallyOutsideStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyStructuredReferenceRewrite.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorksheet("Source");
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

                document.CopyWorksheet(source, "Copy");
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
        public void Test_CopyWorksheetFrom_CopiesValuesBetweenWorkbooks() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopySource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(2, 2, "Region");
                source.CellValue(2, 3, "Revenue");
                source.CellValue(3, 2, "NA");
                source.CellValue(3, 3, 125.5m);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet copied = targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported");

                Assert.Equal("Imported", copied.Name);
                Assert.Equal("B2:C3", copied.GetUsedRangeA1());
                targetDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_CopyWorksheetFrom_PreservesTablesBetweenWorkbooks() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyTableSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyTableTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "NA");
                source.CellValue(2, 2, 100);
                source.CellValue(3, 1, "EMEA");
                source.CellValue(3, 2, 200);
                source.AddTable("A1:B3", hasHeader: true, name: "SourceSales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet copied = targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported");

                Assert.Equal("A1:B3", copied.GetTableRange("SourceSales"));
                targetDocument.Save();
            }

            using (var document = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_CopyWorksheetFrom_PackageModePreservesHeaderOnlyStyles() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageHeaderSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageHeaderTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Headers");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellBold(1, 1, true);
                source.CellBold(1, 2, true);
                source.CellBackground(1, 1, "#D9EAD3");
                source.CellBackground(1, 2, "#D9EAD3");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Summary").CellValue(1, 1, "Summary");
                ExcelSheet copied = targetDocument.CopyWorksheetFrom(sourceDocument, "Headers", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });

                Assert.Equal("Imported", copied.Name);
                Assert.Equal("A1:B1", copied.GetUsedRangeA1());
                targetDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_CopyWorksheetFrom_PackageModeDataTableReadsInlineStrings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageInlineSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageInlineTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("External");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Imported");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "External", "ExternalCopy", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
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
        public void Test_CopyWorksheetFrom_PackageModeMapsSourceDefaultStyleToUnstyledCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDefaultStyleSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDefaultStyleTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("StyledDefault");
                source.CellValue(1, 1, "Default styled");
                sourceDocument.Save();
            }

            AddDefaultBoldFontStyle(sourcePath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellValue(1, 1, "Existing");
                targetDocument.CopyWorksheetFrom(sourceDocument, "StyledDefault", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell copiedCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Assert.NotNull(copiedCell.StyleIndex);

                Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                CellFormat copiedFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)copiedCell.StyleIndex!.Value);
                Font copiedFont = stylesheet.Fonts!.Elements<Font>().ElementAt((int)copiedFormat.FontId!.Value);
                Assert.NotNull(copiedFont.Bold);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModePreservesRowAndColumnStyleInheritance() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageInheritedStyleSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageInheritedStyleTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Inherited");
                source.CellValue(1, 1, "Row inherited");
                source.CellValue(2, 2, "Column inherited");
                sourceDocument.Save();
            }

            AddDefaultBoldFontStyle(sourcePath);
            AddSourceRowAndColumnStyles(sourcePath, "Inherited");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellValue(1, 1, "Existing");
                targetDocument.CopyWorksheetFrom(sourceDocument, "Inherited", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell rowInheritedCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Cell columnInheritedCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B2");
                Row styledRow = copiedPart.Worksheet.Descendants<Row>().Single(row => row.RowIndex?.Value == 1U);
                Column styledColumn = copiedPart.Worksheet.Elements<Columns>().Single().Elements<Column>().Single(column => column.Min?.Value == 2U);

                Assert.Null(rowInheritedCell.StyleIndex);
                Assert.Null(columnInheritedCell.StyleIndex);
                Assert.NotNull(styledRow.StyleIndex);
                Assert.NotNull(styledColumn.Style);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeStreamSavePersistsCopiedSheet() {
            using var sourceStream = new MemoryStream();
            using var targetSeedStream = new MemoryStream();
            using var savedStream = new MemoryStream();

            using (var sourceDocument = ExcelDocument.Create(sourceStream)) {
                sourceDocument.AddWorksheet("Source").CellValue(1, 1, "Copied");
                sourceDocument.Save(sourceStream);
            }

            using (var targetDocument = ExcelDocument.Create(targetSeedStream)) {
                targetDocument.AddWorksheet("Existing").CellValue(1, 1, "Existing");
                targetDocument.Save(targetSeedStream);
            }

            sourceStream.Position = 0;
            targetSeedStream.Position = 0;
            using (var sourceDocument = ExcelDocument.Load(sourceStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetSeedStream)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save(savedStream);
            }

            savedStream.Position = 0;
            using var reloaded = ExcelDocument.Load(savedStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal(2, reloaded.Sheets.Count);
            Assert.True(reloaded["Imported"].TryGetCellText(1, 1, out var value));
            Assert.Equal("Copied", value);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRemapsStylesAndConditionalFormatDxf() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageStyleSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageStyleTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Styled");
                source.CellValue(1, 1, 10);
                source.CellBackground(1, 1, "#D9EAD3");
                sourceDocument.Save();
            }

            AddSourceWorksheetPackageArtifacts(sourcePath, "Styled");

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellValue(1, 1, "Existing");
                targetDocument.Save();
            }

            AddDummyDifferentialFormat(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Styled", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Worksheet worksheet = copiedPart.Worksheet;
                Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                uint cellFormatCount = (uint)stylesheet.CellFormats!.Elements<CellFormat>().Count();
                uint differentialFormatCount = (uint)stylesheet.DifferentialFormats!.Elements<DifferentialFormat>().Count();

                Row row = Assert.Single(worksheet.Descendants<Row>(), item => item.RowIndex?.Value == 1U);
                Column column = Assert.Single(worksheet.Elements<Columns>().Single().Elements<Column>());
                ConditionalFormattingRule rule = Assert.Single(worksheet.Descendants<ConditionalFormattingRule>());

                Assert.True(row.StyleIndex?.Value > 0U && row.StyleIndex.Value < cellFormatCount);
                Assert.True(column.Style?.Value > 0U && column.Style.Value < cellFormatCount);
                Assert.True(rule.FormatId?.Value == 1U && rule.FormatId.Value < differentialFormatCount);
                Assert.Empty(worksheet.Elements<OleObjects>());
                Assert.Empty(worksheet.Elements<Controls>());
                Assert.Empty(worksheet.Elements<Picture>());
                Assert.Empty(worksheet.Elements<LegacyDrawing>());
                Assert.Null(worksheet.GetFirstChild<PageSetup>()?.Id?.Value);
                Assert.DoesNotContain(worksheet.Descendants<OpenXmlElement>(), element => element.LocalName == "queryTableParts");
                Assert.DoesNotContain(worksheet.Descendants<OpenXmlElement>(), element => element.LocalName == "pivotTableDefinition");
                Assert.DoesNotContain(worksheet.Descendants<OpenXmlElement>(), element => element.LocalName == "customProperties");
                Assert.DoesNotContain(worksheet.Descendants<OpenXmlElement>(), element => element.LocalName == "slicerList");
                Assert.DoesNotContain(worksheet.Descendants<OpenXmlElement>(), element => element.LocalName == "timelineRefs");

                OpenXmlElement extensionRule = Assert.Single(worksheet.Descendants<OpenXmlElement>(),
                    element => element.LocalName == "cfRule" && element.NamespaceUri.IndexOf("spreadsheetml/2009", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.Equal("1", extensionRule.GetAttribute("dxfId", string.Empty).Value);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRewritesSelfReferencesAndNamedRanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageFormulaSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageFormulaTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, 10);
                source.CellFormula(2, 1, "Source!A1*TaxRate");
                sourceDocument.SetNamedRange("TaxRate", "A1", source);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
                DefinedName definedName = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>(), name => name.Name == "TaxRate");

                Assert.Equal("'Imported'!A1*TaxRate", formulaCell.CellFormula?.Text);
                Assert.Equal("'Imported'!$A$1", definedName.Text);
                Assert.NotNull(definedName.LocalSheetId);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeMaterializesDeferredDirectExports() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDeferredSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDeferredTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                var table = new DataTable();
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Count", typeof(int));
                table.Rows.Add("Alpha", 1);
                table.Rows.Add("Beta", 2);
                source.InsertDataTableAsTable(table, tableName: "Items");

                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(targetDocument["Imported"].TryGetCellText(1, 1, out var header));
                Assert.True(targetDocument["Imported"].TryGetCellText(3, 1, out var value));
                Assert.Equal("Name", header);
                Assert.Equal("Beta", value);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRewritesStructuredReferencesInWorksheetFormulas() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageStructuredSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageStructuredTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Ada");
                source.AddTable("A1:A2", hasHeader: true, name: "People", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellValue(1, 3, "People");
                source.CellValue(1, 4, "Amount");
                source.CellValue(2, 3, "Ada");
                source.CellValue(2, 4, 10);
                source.AddTable("C1:D2", hasHeader: true, name: "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.ValidationCustomFormula("B2", "COUNTIF(People[Name],B2)>0");
                source.CellFormula(3, 1, "ROWS(People)");
                source.CellFormula(4, 1, "'People'!A1+ROWS(People)");
                source.CellFormula(5, 1, "SUM(Sales[People])+ROWS(People)");
                source.CellFormula(6, 1, "People.Total+ROWS(People)");
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet existing = targetDocument.AddWorksheet("Existing");
                existing.CellValue(1, 1, "Name");
                existing.CellValue(2, 1, "Grace");
                existing.AddTable("A1:A2", hasHeader: true, name: "People", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                targetDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                string copiedTableName = copiedPart.TableDefinitionParts
                    .Select(part => part.Table.Name!.Value!)
                    .Single(name => !string.Equals(name, "Sales", StringComparison.OrdinalIgnoreCase));
                Assert.NotEqual("People", copiedTableName);
                Formula1 formula = Assert.Single(copiedPart.Worksheet.Descendants<Formula1>());
                Cell bareTableFormula = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A3");
                Cell sheetQualifiedFormula = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A4");
                Cell bracketedColumnFormula = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A5");
                Cell dottedNameFormula = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A6");
                Assert.Equal($"COUNTIF({copiedTableName}[Name],B2)>0", formula.Text);
                Assert.Equal($"ROWS({copiedTableName})", bareTableFormula.CellFormula?.Text);
                Assert.Equal($"'People'!A1+ROWS({copiedTableName})", sheetQualifiedFormula.CellFormula?.Text);
                Assert.Equal($"SUM(Sales[People])+ROWS({copiedTableName})", bracketedColumnFormula.CellFormula?.Text);
                Assert.Equal($"People.Total+ROWS({copiedTableName})", dottedNameFormula.CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeDoesNotRewriteFunctionCallsAsTableReferences() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageFunctionNameTableSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageFunctionNameTableTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Value");
                source.CellValue(2, 1, 10);
                source.AddTable("A1:A2", hasHeader: true, name: "SUM", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellFormula(4, 1, "SUM(A2:A2)+ROWS(SUM)");
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet existing = targetDocument.AddWorksheet("Existing");
                existing.CellValue(1, 1, "Value");
                existing.CellValue(2, 1, 20);
                existing.AddTable("A1:A2", hasHeader: true, name: "SUM", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                targetDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                string copiedTableName = Assert.Single(copiedPart.TableDefinitionParts).Table.Name!.Value!;
                Assert.NotEqual("SUM", copiedTableName);
                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A4");
                Assert.Equal($"SUM(A2:A2)+ROWS({copiedTableName})", formulaCell.CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRewritesTableDisplayNameReferences() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDisplayNameSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDisplayNameTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Amount");
                source.CellValue(2, 1, 10);
                source.AddTable("A1:A2", hasHeader: true, name: "Internal", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellFormula(4, 1, "SUM(Visible[Amount])");
                sourceDocument.Save();
            }

            SetTableDisplayName(sourcePath, "Source", "Internal", "Visible");

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet existing = targetDocument.AddWorksheet("Existing");
                existing.CellValue(1, 1, "Amount");
                existing.CellValue(2, 1, 20);
                existing.AddTable("A1:A2", hasHeader: true, name: "Internal", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                targetDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Table copiedTable = Assert.Single(copiedPart.TableDefinitionParts).Table;
                string copiedTableName = copiedTable.Name!.Value!;
                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A4");

                Assert.NotEqual("Internal", copiedTableName);
                Assert.Equal(copiedTableName, copiedTable.DisplayName?.Value);
                Assert.Equal($"SUM({copiedTableName}[Amount])", formulaCell.CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRewritesStructuredReferencesOnce() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageStructuredOnceSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageStructuredOnceTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Amount");
                source.CellValue(2, 1, 10);
                source.AddTable("A1:A2", hasHeader: true, name: "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellValue(4, 1, "Amount");
                source.CellValue(5, 1, 20);
                source.AddTable("A4:A5", hasHeader: true, name: "Sales2", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellFormula(7, 1, "SUM(Sales[Amount])");
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet existing = targetDocument.AddWorksheet("Existing");
                existing.CellValue(1, 1, "Amount");
                existing.CellValue(2, 1, 1);
                existing.AddTable("A1:A2", hasHeader: true, name: "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                targetDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                string[] copiedTableNames = copiedPart.TableDefinitionParts
                    .Select(part => part.Table.Name!.Value!)
                    .OrderBy(name => name)
                    .ToArray();
                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A7");

                Assert.Equal(new[] { "Sales2", "Sales22" }, copiedTableNames);
                Assert.Equal("SUM(Sales2[Amount])", formulaCell.CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeResolvesThemeAndIndexedStyleColors() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageThemeColorSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageThemeColorTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("ThemeStyled");
                source.CellValue(1, 1, "Theme styled");
                sourceDocument.Save();
            }

            AddThemeAndIndexedStyle(sourcePath, "ThemeStyled");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellValue(1, 1, "Existing");
                targetDocument.CopyWorksheetFrom(sourceDocument, "ThemeStyled", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell copiedCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                CellFormat copiedFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)copiedCell.StyleIndex!.Value);
                Font copiedFont = stylesheet.Fonts!.Elements<Font>().ElementAt((int)copiedFormat.FontId!.Value);
                Fill copiedFill = stylesheet.Fills!.Elements<Fill>().ElementAt((int)copiedFormat.FillId!.Value);
                Color fontColor = copiedFont.Color!;
                ForegroundColor fillColor = copiedFill.PatternFill!.ForegroundColor!;

                Assert.NotNull(fontColor.Rgb);
                Assert.Null(fontColor.Theme);
                Assert.NotNull(fillColor.Rgb);
                Assert.Null(fillColor.Indexed);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRemovesStaleCalculationChainForCopiedFormulas() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageCalcChainSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageCalcChainTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("FormulaSource");
                source.CellValue(1, 1, 10);
                source.CellFormula(2, 1, "A1*2");
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellValue(1, 1, "Existing");
                targetDocument.Save();
            }

            AddDummyCalculationChain(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "FormulaSource", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                Assert.Null(spreadsheet.WorkbookPart!.CalculationChainPart);
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
                Assert.Equal("A1*2", formulaCell.CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_MergeWorkbookFrom_SameWorkbookRewritesCopiedTableReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "WorkbookMergeSameWorkbookTables.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "Amount");
                source.CellValue(2, 1, "Alpha");
                source.CellValue(2, 2, 10);
                source.AddTable("A1:B2", hasHeader: true, name: "People", OfficeIMO.Excel.TableStyle.TableStyleMedium9);

                ExcelSheet summary = document.AddWorksheet("Summary");
                summary.CellFormula(1, 1, "SUM(People[Amount])");

                document.MergeWorkbookFrom(document, new ExcelWorkbookMergeOptions {
                    SheetNames = new[] { "Source", "Summary" },
                    SheetNamePrefix = "Copy_",
                    CopyMode = ExcelWorksheetCopyMode.Package
                });

                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            WorksheetPart copiedSummary = GetWorksheetPartByNameForOperations(spreadsheet, "Copy_Summary");
            Cell formulaCell = copiedSummary.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
            Assert.Equal("SUM(People2[Amount])", formulaCell.CellFormula?.Text);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeConvertsDateSerialsAcrossDateSystems() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDateSystemSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDateSystemTarget.xlsx");
            var date = new DateTime(2026, 6, 23);

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                sourceDocument.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet source = sourceDocument.AddWorksheet("Dates");
                source.CellValue(1, 1, "Created");
                source.CellValue(2, 1, date);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Dates", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false);
            WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
            Cell dateCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
            Assert.Equal(date.ToOADate(), double.Parse(dateCell.CellValue!.Text, CultureInfo.InvariantCulture), 6);

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeConvertsDateSerialsWithInheritedRowStyle() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDateSystemInheritedRowSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDateSystemInheritedRowTarget.xlsx");
            var date = new DateTime(2026, 6, 23);

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                sourceDocument.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet source = sourceDocument.AddWorksheet("Dates");
                source.CellValue(1, 1, "Created");
                source.CellValue(2, 1, date);
                sourceDocument.Save();
            }

            MoveCellStyleToRow(sourcePath, "Dates", "A2");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Dates", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false);
            WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
            Cell dateCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
            Row row = copiedPart.Worksheet.Descendants<Row>().Single(item => item.RowIndex?.Value == 2U);
            Assert.Null(dateCell.StyleIndex);
            Assert.NotNull(row.StyleIndex);
            Assert.Equal(date.ToOADate(), double.Parse(dateCell.CellValue!.Text, CultureInfo.InvariantCulture), 6);

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeDoesNotShiftTimeOnlySerialsAcrossDateSystems() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageTimeOnlySource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageTimeOnlyTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                sourceDocument.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet source = sourceDocument.AddWorksheet("Times");
                source.CellValue(1, 1, "Elapsed");
                source.CellValue(2, 1, TimeSpan.FromHours(12));
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Times", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false);
            WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
            Cell timeCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
            Assert.Equal(0.5D, double.Parse(timeCell.CellValue!.Text, CultureInfo.InvariantCulture), 6);

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeConvertsMonthTimeDateSerialsAcrossDateSystems() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageMonthTimeDateSystemSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageMonthTimeDateSystemTarget.xlsx");
            var date = new DateTime(2026, 6, 23, 14, 30, 0);

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                sourceDocument.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet source = sourceDocument.AddWorksheet("Dates");
                source.CellValue(1, 1, "Created");
                source.CellValue(2, 1, date);
                source.CellAt(2, 1).SetNumberFormat("mmm h:mm");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Dates", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false);
            WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
            Cell dateCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
            Assert.Equal(date.ToOADate(), double.Parse(dateCell.CellValue!.Text, CultureInfo.InvariantCulture), 6);

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeDoesNotShiftBracketedElapsedMinuteSerialsAcrossDateSystems() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageElapsedMinuteSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageElapsedMinuteTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                sourceDocument.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet source = sourceDocument.AddWorksheet("Elapsed");
                source.CellValue(1, 1, "Duration");
                source.CellValue(2, 1, 1.25D);
                source.CellAt(2, 1).SetNumberFormat("[m]:ss");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Elapsed", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false);
            WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
            Cell durationCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
            Assert.Equal(1.25D, double.Parse(durationCell.CellValue!.Text, CultureInfo.InvariantCulture), 6);

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeCopiesExternalWorkbookReferences() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageExternalLinkSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageExternalLinkTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("External");
                source.CellFormula(1, 1, "[1]Sheet1!A1");
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellFormula(1, 1, "[1]Sheet1!B1");
                targetDocument.Save();
            }

            AddExternalWorkbookReference(sourcePath);
            AddExternalWorkbookReference(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "External", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    CopyExternalWorkbookReferences = true
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell copiedFormula = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Assert.Equal("[2]Sheet1!A1", copiedFormula.CellFormula?.Text);
                Assert.Equal(2, spreadsheet.WorkbookPart!.Workbook.ExternalReferences!.Elements<ExternalReference>().Count());
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRemapsExternalWorkbookReferencesInDefinedNames() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageExternalDefinedNameSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageExternalDefinedNameTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("External");
                source.CellFormula(1, 1, "ExternalValue");
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellFormula(1, 1, "[1]Sheet1!B1");
                targetDocument.Save();
            }

            AddExternalWorkbookReference(sourcePath);
            AddDefinedName(sourcePath, "ExternalValue", "[1]Sheet1!A1");
            AddExternalWorkbookReference(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "External", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    CopyExternalWorkbookReferences = true
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                DefinedName copiedName = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>(),
                    name => string.Equals(name.Name?.Value, "ExternalValue", StringComparison.OrdinalIgnoreCase));
                Assert.Equal("[2]Sheet1!A1", copiedName.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeDoesNotRemapNumericStructuredReferenceColumnsAsExternalLinks() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageNumericStructuredColumnSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageNumericStructuredColumnTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Structured");
                source.CellValue(1, 1, "1");
                source.CellValue(2, 1, 7);
                source.AddTable("A1:A2", hasHeader: true, name: "People", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                source.CellFormula(4, 1, "SUM(People[1])");
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet existing = targetDocument.AddWorksheet("Existing");
                existing.CellValue(1, 1, "1");
                existing.CellValue(2, 1, 3);
                existing.AddTable("A1:A2", hasHeader: true, name: "People", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                existing.CellFormula(4, 1, "[1]Sheet1!B1");
                targetDocument.Save();
            }

            AddExternalWorkbookReference(sourcePath);
            AddExternalWorkbookReference(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Structured", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    CopyExternalWorkbookReferences = true
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell copiedFormula = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A4");
                Assert.Equal("SUM(People2[1])", copiedFormula.CellFormula?.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeStripsCopiedTableQueryBindings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageQueryTableSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageQueryTableTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Query");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Ada");
                source.AddTable("A1:A2", hasHeader: true, name: "QueryTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                sourceDocument.Save();
            }

            AddTableQueryBindings(sourcePath, "Query");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Query", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Table table = Assert.Single(copiedPart.TableDefinitionParts).Table!;
                Assert.Null(table.ConnectionId);
                Assert.All(table.Descendants<TableColumn>(), column => Assert.Null(column.QueryTableFieldId));
                Assert.DoesNotContain(table.Descendants<OpenXmlElement>(), element =>
                    string.Equals(element.LocalName, "queryTable", StringComparison.Ordinal)
                    || string.Equals(element.LocalName, "queryTableField", StringComparison.Ordinal)
                    || string.Equals(element.LocalName, "queryTableFields", StringComparison.Ordinal));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeSavesTableFormulaExternalReferenceRemaps() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageTableExternalFormulaSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageTableExternalFormulaTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("TableFormula");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "External");
                source.CellValue(2, 1, "Ada");
                source.CellValue(2, 2, 0);
                source.AddTable("A1:B2", hasHeader: true, name: "ExternalTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                sourceDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddWorksheet("Existing").CellFormula(1, 1, "[1]Sheet1!B1");
                targetDocument.Save();
            }

            AddExternalWorkbookReference(sourcePath);
            AddTableCalculatedColumnFormula(sourcePath, "TableFormula", "External", "[1]Sheet1!A1");
            AddExternalWorkbookReference(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "TableFormula", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    CopyExternalWorkbookReferences = true
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Table table = Assert.Single(copiedPart.TableDefinitionParts).Table!;
                CalculatedColumnFormula formula = Assert.Single(table.Descendants<CalculatedColumnFormula>());
                Assert.Equal("[2]Sheet1!A1", formula.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeMaterializesTargetDeferredImportsBeforeCopy() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDeferredTargetSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageDeferredTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                sourceDocument.AddWorksheet("Source").CellValue(1, 1, "Copied");
                sourceDocument.Save();
            }

            var dataSet = new DataSet("Deferred");
            DataTable table = dataSet.Tables.Add("Pending");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Ada");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.InsertDataSet(dataSet, createTables: true, autoFit: false);
                targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (var reloaded = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(reloaded["Pending"].TryGetCellText(2, 1, out var pendingValue));
                Assert.Equal("Ada", pendingValue);
                Assert.True(reloaded["Imported"].TryGetCellText(1, 1, out var copiedValue));
                Assert.Equal("Copied", copiedValue);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeStripsCellMetadataIndexes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageMetadataSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyPackageMetadataTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Metadata");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Alpha");
                sourceDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(sourcePath, true)) {
                WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, "Metadata");
                Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "A2");
                cell.SetAttribute(new OpenXmlAttribute(string.Empty, "cm", string.Empty, "1"));
                cell.SetAttribute(new OpenXmlAttribute(string.Empty, "vm", string.Empty, "2"));
                worksheetPart.Worksheet.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Metadata", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell copiedCell = copiedPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "A2");
                Assert.DoesNotContain(copiedCell.GetAttributes(), attribute => attribute.LocalName == "cm" || attribute.LocalName == "vm");
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_ValuesModeHonorsSameWorkbookSource() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCopyValuesModeSameWorkbook.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Ada");
                source.CellBold(1, 1, true);
                ExcelSheet copied = document.CopyWorksheetFrom(document, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Values
                });

                Assert.Equal("Imported", copied.Name);
                Assert.Equal("A1:A2", copied.GetUsedRangeA1());
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(document["Imported"].TryGetCellText(2, 1, out var value));
                Assert.Equal("Ada", value);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
                Cell headerCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Assert.True(headerCell.StyleIndex == null || headerCell.StyleIndex.Value == 0U);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_ValuesModeKeepsReaderWriterFallback() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyValuesModeSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyValuesModeTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(2, 1, "Ada");
                source.CellBold(1, 1, true);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet copied = targetDocument.CopyWorksheetFrom(sourceDocument, "Source", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Values
                });

                Assert.Equal("Imported", copied.Name);
                Assert.Equal("A1:A2", copied.GetUsedRangeA1());
                targetDocument.Save();
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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

        private static void AddSourceWorksheetPackageArtifacts(string path, string sheetName) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Worksheet worksheet = worksheetPart.Worksheet;
            Cell cell = worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "A1");
            uint styleIndex = cell.StyleIndex?.Value ?? 0U;

            SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
            Row row = sheetData.Elements<Row>().Single(item => item.RowIndex?.Value == 1U);
            row.StyleIndex = styleIndex;
            row.CustomFormat = true;
            worksheet.InsertBefore(new Columns(new Column {
                Min = 1U,
                Max = 1U,
                Style = styleIndex,
                Width = 14D,
                CustomWidth = true
            }), sheetData);

            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            stylesheet.DifferentialFormats ??= new DifferentialFormats();
            stylesheet.DifferentialFormats.Append(new DifferentialFormat(new Fill(new PatternFill(new ForegroundColor { Rgb = "FFFF0000" }) {
                PatternType = PatternValues.Solid
            })));
            stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.Elements<DifferentialFormat>().Count();
            stylesheet.Save();

            worksheet.Append(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("A1>0")) {
                    Type = ConditionalFormatValues.Expression,
                    FormatId = 0U,
                    Priority = 1
                }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            });
            worksheet.Append(new OleObjects());
            worksheet.Append(new Controls());
            worksheet.Append(new Picture { Id = "rId999" });
            worksheet.Append(new LegacyDrawing { Id = "rId998" });
            worksheet.Append(new PageSetup { Id = "rId997" });
            var customProperties = new OpenXmlUnknownElement(string.Empty, "customProperties", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var customProperty = new OpenXmlUnknownElement(string.Empty, "customPr", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            customProperty.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId994"));
            customProperties.Append(customProperty);
            worksheet.Append(customProperties);
            var queryTableParts = new OpenXmlUnknownElement(string.Empty, "queryTableParts", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var queryTablePart = new OpenXmlUnknownElement(string.Empty, "queryTablePart", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            queryTablePart.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId996"));
            queryTableParts.Append(queryTablePart);
            worksheet.Append(queryTableParts);
            var pivotTableDefinition = new OpenXmlUnknownElement(string.Empty, "pivotTableDefinition", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            pivotTableDefinition.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId995"));
            worksheet.Append(pivotTableDefinition);

            var extensionList = worksheet.GetFirstChild<WorksheetExtensionList>() ?? worksheet.AppendChild(new WorksheetExtensionList());
            var slicerExtension = new WorksheetExtension { Uri = "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" };
            var slicerList = new OpenXmlUnknownElement("x14", "slicerList", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var slicer = new OpenXmlUnknownElement("x14", "slicer", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            slicer.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId993"));
            slicerList.Append(slicer);
            slicerExtension.Append(slicerList);
            extensionList.Append(slicerExtension);

            var timelineExtension = new WorksheetExtension { Uri = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}" };
            var timelineRefs = new OpenXmlUnknownElement("x15", "timelineRefs", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            var timelineRef = new OpenXmlUnknownElement("x15", "timelineRef", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            timelineRef.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId992"));
            timelineRefs.Append(timelineRef);
            timelineExtension.Append(timelineRefs);
            extensionList.Append(timelineExtension);

            var conditionalExtension = new WorksheetExtension { Uri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}" };
            var conditionalFormattings = new OpenXmlUnknownElement("x14", "conditionalFormattings", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var conditionalFormatting = new OpenXmlUnknownElement("x14", "conditionalFormatting", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var extensionRule = new OpenXmlUnknownElement("x14", "cfRule", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            extensionRule.SetAttribute(new OpenXmlAttribute(string.Empty, "type", string.Empty, "expression"));
            extensionRule.SetAttribute(new OpenXmlAttribute(string.Empty, "dxfId", string.Empty, "0"));
            conditionalFormatting.Append(extensionRule);
            conditionalFormattings.Append(conditionalFormatting);
            conditionalExtension.Append(conditionalFormattings);
            extensionList.Append(conditionalExtension);
            worksheet.Save();
        }

        private static void AddDefaultBoldFontStyle(string path) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            stylesheet.Fonts ??= new Fonts();
            stylesheet.Fonts.Append(new Font(new Bold()));
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Elements<Font>().Count();

            CellFormat defaultFormat = stylesheet.CellFormats!.Elements<CellFormat>().First();
            defaultFormat.FontId = stylesheet.Fonts.Count!.Value - 1U;
            defaultFormat.ApplyFont = true;
            stylesheet.Save();
        }

        private static void AddSourceRowAndColumnStyles(string path, string sheetName) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;

            stylesheet.Fonts ??= new Fonts();
            stylesheet.Fonts.Append(new Font(new Italic()));
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Elements<Font>().Count();
            uint fontId = stylesheet.Fonts.Count!.Value - 1U;

            stylesheet.CellFormats ??= new CellFormats();
            stylesheet.CellFormats.Append(new CellFormat {
                FontId = fontId,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyFont = true
            });
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
            uint styleIndex = stylesheet.CellFormats.Count!.Value - 1U;

            Row row = sheetData.Elements<Row>().Single(item => item.RowIndex?.Value == 1U);
            row.StyleIndex = styleIndex;
            row.CustomFormat = true;
            worksheet.InsertBefore(new Columns(new Column {
                Min = 2U,
                Max = 2U,
                Style = styleIndex,
                Width = 14D,
                CustomWidth = true
            }), sheetData);

            stylesheet.Save();
            worksheet.Save();
        }

        private static void AddDummyDifferentialFormat(string path) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            stylesheet.DifferentialFormats ??= new DifferentialFormats();
            stylesheet.DifferentialFormats.Append(new DifferentialFormat(new Font(new Bold())));
            stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.Elements<DifferentialFormat>().Count();
            stylesheet.Save();
        }

        private static void AddThemeAndIndexedStyle(string path, string sheetName) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Worksheet worksheet = worksheetPart.Worksheet;
            Cell cell = worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "A1");
            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;

            stylesheet.Fonts ??= new Fonts();
            stylesheet.Fonts.Append(new Font(new Color { Theme = 4U }));
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Elements<Font>().Count();
            uint fontId = stylesheet.Fonts.Count!.Value - 1U;

            stylesheet.Fills ??= new Fills();
            stylesheet.Fills.Append(new Fill(new PatternFill(new ForegroundColor { Indexed = 10U }) {
                PatternType = PatternValues.Solid
            }));
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Elements<Fill>().Count();
            uint fillId = stylesheet.Fills.Count!.Value - 1U;

            stylesheet.CellFormats ??= new CellFormats();
            stylesheet.CellFormats.Append(new CellFormat {
                FontId = fontId,
                FillId = fillId,
                BorderId = 0U,
                FormatId = 0U,
                ApplyFont = true,
                ApplyFill = true
            });
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();

            cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
            stylesheet.Save();
            worksheet.Save();
        }

        private static void AddDummyCalculationChain(string path) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            if (workbookPart.CalculationChainPart != null) {
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
            }

            CalculationChainPart calculationChainPart = workbookPart.AddNewPart<CalculationChainPart>();
            calculationChainPart.CalculationChain = new CalculationChain(
                new CalculationCell {
                    CellReference = "A1",
                    SheetId = 1
                });
            calculationChainPart.CalculationChain.Save();
        }

        private static void AddExternalWorkbookReference(string path) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            ExternalWorkbookPart externalPart = workbookPart.AddNewPart<ExternalWorkbookPart>();
            externalPart.ExternalLink = new ExternalLink(
                new ExternalBook(
                    new SheetNames(
                        new SheetName { Val = "Sheet1" })));
            externalPart.ExternalLink.Save();

            string relationshipId = workbookPart.GetIdOfPart(externalPart);
            workbookPart.Workbook.ExternalReferences ??= new ExternalReferences();
            workbookPart.Workbook.ExternalReferences.Append(new ExternalReference { Id = relationshipId });
            workbookPart.Workbook.Save();
        }

        private static void AddDefinedName(string path, string name, string formula, uint? localSheetId = null) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
            workbook.DefinedNames ??= new DefinedNames();
            workbook.DefinedNames.Append(new DefinedName(formula) {
                Name = name,
                LocalSheetId = localSheetId
            });
            workbook.Save();
        }

        private static void AddTableQueryBindings(string path, string sheetName) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Table table = worksheetPart.TableDefinitionParts.Single().Table!;
            table.ConnectionId = 1U;
            TableColumn column = table.Descendants<TableColumn>().Single();
            column.QueryTableFieldId = 1U;

            var extensionList = table.GetFirstChild<TableExtensionList>() ?? table.AppendChild(new TableExtensionList());
            var extension = new TableExtension { Uri = "{00000000-0000-0000-0000-000000000001}" };
            var queryTableFields = new OpenXmlUnknownElement("x14", "queryTableFields", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var queryTableField = new OpenXmlUnknownElement("x14", "queryTableField", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            queryTableFields.Append(queryTableField);
            extension.Append(queryTableFields);
            extensionList.Append(extension);
            table.Save();
        }

        private static void AddTableCalculatedColumnFormula(string path, string sheetName, string columnName, string formula) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Table table = worksheetPart.TableDefinitionParts.Single().Table!;
            TableColumn column = table.Descendants<TableColumn>()
                .Single(item => string.Equals(item.Name?.Value, columnName, StringComparison.OrdinalIgnoreCase));
            column.RemoveAllChildren<CalculatedColumnFormula>();
            column.Append(new CalculatedColumnFormula(formula));
            table.Save();
        }

        [Fact]
        public void Test_CompareRanges_ReturnsCellDifferences() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetCompare.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet left = document.AddWorksheet("Left");
                left.CellValue(1, 1, "Name");
                left.CellValue(1, 2, "Score");
                left.CellValue(2, 1, "Ada");
                left.CellValue(2, 2, 10);

                ExcelSheet right = document.AddWorksheet("Right");
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
        public void Test_MergeWorksheets_AppendsRowsAndSkipsSourceHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeAppend.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorksheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");
                target.CellValue(2, 1, "NA");
                target.CellValue(2, 2, 100);

                ExcelSheet source = document.AddWorksheet("More");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "EMEA");
                source.CellValue(2, 2, 200);
                source.CellValue(3, 1, "APAC");
                source.CellValue(3, 2, 150);

                ExcelWorksheetMergeResult result = document.MergeWorksheets(target, source);

                Assert.Equal("Combined", result.TargetSheetName);
                Assert.Equal("More", result.SourceSheetName);
                Assert.Equal("A1:B3", result.SourceRange);
                Assert.Equal("A3:B4", result.TargetRange);
                Assert.Equal(2, result.RowsCopied);
                Assert.Equal(2, result.ColumnsCopied);
                Assert.True(result.HeaderSkipped);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_MergeWorksheets_CanMatchColumnsByHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeHeaders.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorksheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");
                target.CellValue(2, 1, "NA");
                target.CellValue(2, 2, 100);

                ExcelSheet source = document.AddWorksheet("More");
                source.CellValue(1, 1, "Revenue");
                source.CellValue(1, 2, "Region");
                source.CellValue(2, 1, 200);
                source.CellValue(2, 2, "EMEA");
                source.CellValue(3, 1, 150);
                source.CellValue(3, 2, "APAC");

                ExcelWorksheetMergeResult result = document.MergeWorksheets(target, source, new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true
                });

                Assert.Equal("A3:B4", result.TargetRange);
                Assert.Equal(2, result.RowsCopied);
                Assert.Equal(2, result.ColumnsCopied);
                Assert.True(result.HeaderSkipped);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_MergeWorksheets_CanMatchColumnsUsingExplicitTargetHeaderRow() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeExplicitHeaderRow.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorksheet("Combined");
                target.CellValue(1, 1, "Quarterly report");
                target.CellValue(3, 2, "Region");
                target.CellValue(3, 3, "Revenue");
                target.CellValue(4, 2, "NA");
                target.CellValue(4, 3, 100);

                ExcelSheet source = document.AddWorksheet("More");
                source.CellValue(1, 1, "Revenue");
                source.CellValue(1, 2, "Region");
                source.CellValue(2, 1, 200);
                source.CellValue(2, 2, "EMEA");

                ExcelWorksheetMergeResult result = document.MergeWorksheets(target, source, new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true,
                    TargetHeaderRow = 3,
                    TargetStartRow = 5,
                    TargetStartColumn = 2
                });

                Assert.Equal("B5:C5", result.TargetRange);
                Assert.Equal(1, result.RowsCopied);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_MergeWorksheets_HeaderMatchThrowsWhenSourceColumnIsMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetMergeHeadersMissing.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorksheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");
                target.CellValue(2, 1, "NA");
                target.CellValue(2, 2, 100);

                ExcelSheet source = document.AddWorksheet("More");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Amount");
                source.CellValue(2, 1, "EMEA");
                source.CellValue(2, 2, 200);

                var exception = Assert.Throws<ArgumentException>(() => document.MergeWorksheets(target, source, new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true
                }));
                Assert.Contains("Revenue", exception.Message);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_MergeWorksheets_AllowsSourceFromAnotherWorkbook() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetMergeExternalSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetMergeExternalTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("More");
                source.CellValue(1, 1, "Revenue");
                source.CellValue(1, 2, "Region");
                source.CellValue(2, 1, 200);
                source.CellValue(2, 2, "EMEA");
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet target = targetDocument.AddWorksheet("Combined");
                target.CellValue(1, 1, "Region");
                target.CellValue(1, 2, "Revenue");

                ExcelWorksheetMergeResult result = targetDocument.MergeWorksheets(target, sourceDocument.GetSheet("More"), new ExcelWorksheetMergeOptions {
                    MatchColumnsByHeader = true
                });

                Assert.Equal("A2:B2", result.TargetRange);
                Assert.Equal(1, result.RowsCopied);
                targetDocument.Save();
            }

            using (var document = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
        public void Test_MergeWorksheets_ThrowsWhenTargetBelongsToAnotherWorkbook() {
            string leftPath = Path.Combine(_directoryWithFiles, "WorksheetMergeWrongTargetLeft.xlsx");
            string rightPath = Path.Combine(_directoryWithFiles, "WorksheetMergeWrongTargetRight.xlsx");

            using (var leftDocument = ExcelDocument.Create(leftPath))
            using (var rightDocument = ExcelDocument.Create(rightPath)) {
                ExcelSheet wrongTarget = rightDocument.AddWorksheet("WrongTarget");
                wrongTarget.CellValue(1, 1, "Region");

                ExcelSheet source = leftDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(2, 1, "EMEA");

                var exception = Assert.Throws<ArgumentException>(() => leftDocument.MergeWorksheets(wrongTarget, source));
                Assert.Contains("Target worksheet", exception.Message);
            }

            File.Delete(leftPath);
            File.Delete(rightPath);
        }

        [Fact]
        public void Test_JoinWorksheets_CopiesRequestedRangeToExplicitPosition() {
            string filePath = Path.Combine(_directoryWithFiles, "WorksheetJoinRange.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet target = document.AddWorksheet("Target");
                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "ignore");
                source.CellValue(2, 2, "Name");
                source.CellValue(2, 3, "Score");
                source.CellValue(3, 2, "Ada");
                source.CellValue(3, 3, 10);

                ExcelWorksheetMergeResult result = document.MergeWorksheets(target, source, new ExcelWorksheetMergeOptions {
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                ExcelSheet target = document.AddWorksheet("Target");
                target.CellValue(5, 4, "Existing");

                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "Score");
                source.CellValue(2, 1, "Ada");
                source.CellValue(2, 2, 10);

                var exception = Assert.Throws<InvalidOperationException>(() => document.MergeWorksheets(target, source, new ExcelWorksheetMergeOptions {
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
                ExcelSheet target = document.AddWorksheet("Target");
                target.CellValue(5, 4, "Existing");

                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Name");
                source.CellValue(1, 2, "Score");
                source.CellValue(2, 1, "Ada");
                source.CellValue(2, 2, 10);

                ExcelWorksheetMergeResult result = document.MergeWorksheets(target, source, new ExcelWorksheetMergeOptions {
                    SourceRange = "A1:B2",
                    IncludeSourceHeader = true,
                    TargetStartRow = 5,
                    TargetStartColumn = 4,
                    OverwriteExistingCells = true
                });

                Assert.Equal("D5:E6", result.TargetRange);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                using var reader = document.CreateReader();
                object?[,] values = reader.GetSheet("Target").ReadRange("D5:E6");
                Assert.Equal("Name", values[0, 0]);
                Assert.Equal("Score", values[0, 1]);
                Assert.Equal("Ada", values[1, 0]);
                Assert.Equal(10d, values[1, 1]);
            }

            File.Delete(filePath);
        }

        private static void SetTableDisplayName(string path, string sheetName, string tableName, string displayName) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Table table = worksheetPart.TableDefinitionParts
                .Select(part => part.Table)
                .Single(table => string.Equals(table?.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase));
            table.DisplayName = displayName;
            table.Save();
        }

        private static void MoveCellStyleToRow(string path, string sheetName, string cellReference) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == cellReference);
            Assert.NotNull(cell.StyleIndex);
            (int rowIndex, _) = A1.ParseCellRef(cellReference);
            Row row = worksheetPart.Worksheet.Descendants<Row>().Single(item => item.RowIndex?.Value == (uint)rowIndex);
            row.StyleIndex = cell.StyleIndex!.Value;
            row.CustomFormat = true;
            cell.StyleIndex = null;
            worksheetPart.Worksheet.Save();
        }

        private static WorksheetPart GetWorksheetPartByNameForOperations(SpreadsheetDocument document, string sheetName) {
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>()
                .Single(candidate => candidate.Name?.Value == sheetName);
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        }
    }
}
