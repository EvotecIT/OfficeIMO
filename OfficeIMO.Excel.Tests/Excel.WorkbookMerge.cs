using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelWorkbookMerge_ImportsSelectedSheetsWithPrefix() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.Source.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.Target.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddWorksheet("North").CellValue(1, 1, "North value");
                source.AddWorksheet("South").CellValue(1, 1, "South value");
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                target.AddWorksheet("Summary");
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    SheetNames = new[] { "South" },
                    SheetNamePrefix = "Imported "
                });

                Assert.Equal(1, result.SheetCount);
                Assert.Equal(new[] { "South" }, result.SourceSheets);
                Assert.Equal(new[] { "Imported South" }, result.TargetSheets);
                Assert.True(target["Imported South"].TryGetCellText(1, 1, out var importedValue));
                Assert.Equal("South value", importedValue);
                target.Save();
            }

            using (var reloaded = ExcelDocument.Load(targetPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(reloaded["Imported South"].TryGetCellText(1, 1, out var importedValue));
                Assert.Equal("South value", importedValue);
            }
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_StreamBackedWorkbookDoesNotForceSave() {
            using var targetStream = new MemoryStream();
            using var sourceStream = new MemoryStream();

            using (var source = ExcelDocument.Create(sourceStream)) {
                source.AddWorksheet("Source").CellValue(1, 1, "Imported");
                source.Save(sourceStream);
            }

            sourceStream.Position = 0;
            using (var target = ExcelDocument.Create(targetStream))
            using (var source = ExcelDocument.Load(sourceStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                target.AddWorksheet("Target");
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(source);

                Assert.Equal(1, result.SheetCount);
                Assert.True(target["Source"].TryGetCellText(1, 1, out var imported));
                Assert.Equal("Imported", imported);
                target.Save(targetStream);
            }

            targetStream.Position = 0;
            using var reloaded = ExcelDocument.Load(targetStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal(2, reloaded.Sheets.Count);
            Assert.True(reloaded["Source"].TryGetCellText(1, 1, out var value));
            Assert.Equal("Imported", value);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_RewritesCopiedWorksheetFormulasForPrefixedNames() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.FormulaSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.FormulaTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet data = source.AddWorksheet("Data");
                data.CellValue(1, 1, "Name");
                data.CellValue(1, 2, "Amount");
                data.CellValue(2, 1, "Ada");
                data.CellValue(2, 2, 42);
                data.CellValue(2, 3, 0.2);
                data.CellFormula(3, 2, "Data!B2");
                data.AddTable("A1:B2", hasHeader: true, name: "People", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.SetNamedRange("TaxRate", "C2", data, save: false);

                ExcelSheet importedData = source.AddWorksheet("Imported Data");
                importedData.CellValue(1, 1, 84);

                ExcelSheet jan = source.AddWorksheet("Jan");
                jan.CellValue(1, 1, 1);

                ExcelSheet mar = source.AddWorksheet("Mar");
                mar.CellValue(1, 1, 3);

                ExcelSheet jan2026 = source.AddWorksheet("Jan 2026");
                jan2026.CellValue(1, 1, 1);

                ExcelSheet mar2026 = source.AddWorksheet("Mar 2026");
                mar2026.CellValue(1, 1, 3);

                ExcelSheet summary = source.AddWorksheet("Summary");
                summary.CellFormula(1, 1, "Data!B2+'Imported Data'!A1+Data!TaxRate+SUM(People[Amount])+TotalWithTax+SUM(Jan:Mar!A1)+SUM('Jan 2026:Mar 2026'!A1)");
                summary.SetInternalLink(2, 1, "Data!A1", "Go");
                summary.ValidationCustomFormula("B2", "COUNTIF(Data!$A$1:$A$1,\">0\")>0");
                source.Save();
            }

            AddWorkbookDefinedName(sourcePath, "TaxRate", "0.99");
            AddWorkbookDefinedName(sourcePath, "PeopleTotal", "SUM(People[Amount])");
            AddWorkbookDefinedName(sourcePath, "TotalWithTax", "PeopleTotal*TaxRate");

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet existing = target.AddWorksheet("Existing");
                existing.CellValue(1, 1, "Name");
                existing.CellValue(2, 1, "Grace");
                existing.AddTable("A1:A2", hasHeader: true, name: "People", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    SheetNamePrefix = "Imported ",
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                target.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart dataPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported Data");
                Cell selfFormulaCell = dataPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B3");
                string copiedPeopleTableName = Assert.Single(dataPart.TableDefinitionParts).Table.Name!.Value!;
                WorksheetPart summaryPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported Summary");
                Cell formulaCell = summaryPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Hyperlink hyperlink = Assert.Single(summaryPart.Worksheet.Descendants<Hyperlink>());
                Formula1 validationFormula = Assert.Single(summaryPart.Worksheet.Descendants<Formula1>());
                DefinedName taxRate = spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>()
                    .Single(name => name.Name == "TaxRate"
                        && name.LocalSheetId != null
                        && name.Text == "'Imported Data'!$C$2");
                DefinedName peopleTotal = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>(), name => name.Name == "PeopleTotal");
                DefinedName totalWithTax = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>(), name => name.Name == "TotalWithTax");

                Assert.NotEqual("People", copiedPeopleTableName);
                Assert.Equal("'Imported Data'!B2", selfFormulaCell.CellFormula?.Text);
                Assert.Equal($"'Imported Data'!B2+'Imported Imported Data'!A1+'Imported Data'!TaxRate+SUM({copiedPeopleTableName}[Amount])+TotalWithTax+SUM('Imported Jan:Imported Mar'!A1)+SUM('Imported Jan 2026:Imported Mar 2026'!A1)", formulaCell.CellFormula?.Text);
                Assert.Equal("'Imported Data'!A1", hyperlink.Location?.Value);
                Assert.Equal("COUNTIF('Imported Data'!$A$1:$A$1,\">0\")>0", validationFormula.Text);
                Assert.Equal("'Imported Data'!$C$2", taxRate.Text);
                Assert.Equal($"SUM({copiedPeopleTableName}[Amount])", peopleTotal.Text);
                Assert.Equal("PeopleTotal*TaxRate", totalWithTax.Text);
                Assert.NotNull(taxRate.LocalSheetId);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_RemapsExternalReferencesInCopiedDefinedNames() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.ExternalNameSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.ExternalNameTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet data = source.AddWorksheet("Data");
                data.CellFormula(1, 1, "ExternalValue");
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath)) {
                target.AddWorksheet("Existing").CellFormula(1, 1, "[1]Sheet1!B1");
                target.Save();
            }

            AddExternalWorkbookReference(sourcePath);
            AddWorkbookDefinedName(sourcePath, "ExternalValue", "[1]Sheet1!A1");
            AddExternalWorkbookReference(targetPath);

            using (var target = ExcelDocument.Load(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    CopyExternalWorkbookReferences = true
                });
                target.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                DefinedName copiedName = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>(),
                    name => string.Equals(name.Name?.Value, "ExternalValue", System.StringComparison.OrdinalIgnoreCase));
                Assert.Equal("[2]Sheet1!A1", copiedName.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_SameWorkbookPackageModeUsesWorksheetCopyPath() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.SameWorkbookPackageMode.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Amount");
                source.CellValue(2, 1, 10);
                source.AddTable("A1:A2", hasHeader: true, name: "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                source.CellFormula(3, 1, "SUM(Sales[Amount])");

                ExcelWorkbookMergeResult result = document.MergeWorkbookFrom(document, new ExcelWorkbookMergeOptions {
                    SheetNames = new[] { "Source" },
                    SheetNamePrefix = "Copy ",
                    CopyMode = ExcelWorksheetCopyMode.Package
                });

                Assert.Equal(new[] { "Source" }, result.SourceSheets);
                Assert.Equal(new[] { "Copy Source" }, result.TargetSheets);
                Assert.Equal("A1:A2", document.GetSheet("Copy Source").GetTableRange("Sales2"));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Copy Source");
                Cell formulaCell = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A3");
                Assert.Equal("SUM(Sales2[Amount])", formulaCell.CellFormula?.Text);
            }

            File.Delete(filePath);
        }

        private static void AddWorkbookDefinedName(string path, string name, string reference) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
            workbook.DefinedNames ??= new DefinedNames();
            workbook.DefinedNames.Append(new DefinedName(reference) {
                Name = name
            });
            workbook.Save();
        }
    }
}
