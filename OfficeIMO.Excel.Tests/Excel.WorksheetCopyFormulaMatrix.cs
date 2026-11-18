using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // One formula exercising every reference kind a copy rewrites: a quoted sheet reference, a structured reference,
        // a chained and a sheet-scoped defined name, an external [n] reference and a name that resolves externally,
        // followed by a string literal holding the same text, which must survive unchanged.
        private const string MatrixLiteral = "\"'My Data'!A1 Sales[Amount] Rate Decoy R1C1 [1]Sheet1!A1\"";
        private const string MatrixFormula =
            "AND('My Data'!$A$1<>\"\",SUM(Sales[Amount])>Chain,Local>Rate,[1]Sheet1!$A$1>0,Ext>0," + MatrixLiteral + "<>\"\")";

        public static IEnumerable<object[]> FormulaMatrixRoutes() {
            foreach (bool merge in new[] { false, true }) {
                foreach (ExcelWorksheetCopyMode mode in new[] { ExcelWorksheetCopyMode.Values, ExcelWorksheetCopyMode.Package }) {
                    yield return new object[] { merge, mode };
                }
            }
        }

        [Theory]
        [MemberData(nameof(FormulaMatrixRoutes))]
        public void Test_WorksheetCopy_FormulaReferenceMatrix(bool merge, ExcelWorksheetCopyMode mode) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"FormulaMatrixSource-{merge}-{mode}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"FormulaMatrixTarget-{merge}-{mode}.xlsx");
            CreateFormulaMatrixSource(sourcePath);

            // The target already has a table named Sales and one external link, so the copied table is renamed and
            // the copied link becomes [2]; the sheet itself is renamed to Copy.
            CreateTableSource(targetPath, "Existing", "Sales");
            AddLinkedWorkbook(targetPath, "Rates.xlsx");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                if (merge) {
                    targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                        CopyMode = mode,
                        SheetNames = new[] { "My Data" },
                        TargetSheetNames = new Dictionary<string, string> { ["My Data"] = "Copy" },
                        CopyExternalWorkbookReferences = true
                    });
                } else {
                    targetDocument.CopyWorksheetFrom(sourceDocument, "My Data", "Copy", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions {
                        CopyMode = mode,
                        CopyExternalWorkbookReferences = true
                    });
                }

                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart copyPart = GetWorksheetPartByNameForOperations(target, "Copy");
                string table = copyPart.TableDefinitionParts.Single().Table!.Name!.Value!;
                Assert.NotEqual("Sales", table);
                string expected = "AND(Copy!$A$1<>\"\",SUM(" + table + "[Amount])>Chain,Local>Rate,[2]Sheet1!$A$1>0,Ext>0," + MatrixLiteral + "<>\"\")";

                Worksheet copy = copyPart.Worksheet!;
                Assert.Equal(expected, Unquote(copy.Descendants<ConditionalFormattingRule>().Single(rule => rule.Type!.Value == ConditionalFormatValues.Expression).GetFirstChild<Formula>()!.Text));
                Assert.Equal(expected, Unquote(copy.Descendants<ConditionalFormatValueObject>().Single(threshold => threshold.Type!.Value == ConditionalFormatValueObjectValues.Formula).Val!.Value!));
                Assert.Equal(expected, Unquote(copy.Descendants<DataValidation>().Single().GetFirstChild<Formula1>()!.Text));

                Workbook workbook = target.WorkbookPart!.Workbook!;
                uint copyPosition = (uint)workbook.Sheets!.Elements<Sheet>().ToList().FindIndex(sheet => sheet.Name == "Copy");
                Dictionary<string, string> names = workbook.DefinedNames!.Elements<DefinedName>()
                    .Where(name => name.LocalSheetId?.Value == copyPosition)
                    .ToDictionary(name => name.Name!.Value!, name => Unquote(name.Text));
                Assert.Equal("0.5", names["Rate"]);
                Assert.Equal("Rate*2", names["Chain"]);
                Assert.Equal("Copy!$B$2", names["Local"]);
                Assert.Equal("[2]Sheet1!$A$1", names["Ext"]);
                Assert.Equal("OFFSET(Copy!$A$1,0,0,ROWS(" + table + "[#All]),Chain)", names["_xlnm.Print_Area"]);
                Assert.Equal("Copy!$1:$1", names["_xlnm.Print_Titles"]);
                // Decoy appears only inside the string literal.
                Assert.DoesNotContain("Decoy", names.Keys);
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            using (var reloaded = ExcelDocument.Load(targetPath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.True(reloaded["Copy"].TryGetCellText(2, 2, out string amount));
                Assert.Equal("10", amount);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_WorksheetValuesCopy_NamesInsideStringLiteralsAreNotDependencies(bool merge) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"LiteralNameSource-{merge}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"LiteralNameTarget-{merge}.xlsx");
            using (var createdSource = ExcelDocument.Create(sourcePath)) {
                createdSource.AddSheet("Data").Cell(1, 1).SetValue("Red");
                createdSource.Save();
            }

            AddLinkedWorkbook(sourcePath, "Prices.xlsx");
            AddDefinedName(sourcePath, "Red", "[1]Sheet1!$A$1");
            // A name spelled like the sheet must not be mistaken for the sheet qualifier in Data!$A$1.
            AddDefinedName(sourcePath, "Data", "[1]Sheet1!$C$1");
            AddDefinedName(sourcePath, "_xlnm.Print_Area", "Data!$A$1:$A$3", localSheetId: 0U);
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                Worksheet worksheet = GetWorksheetPartByNameForOperations(source, "Data").Worksheet!;
                worksheet.InsertAfter(new ConditionalFormatting(new ConditionalFormattingRule(new Formula("Data!$A$1<>\"\"")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1
                }) { SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" } }, worksheet.GetFirstChild<SheetData>()!);
                worksheet.InsertAfter(new DataValidations(new DataValidation(new Formula1("\"Red,Green\"")) {
                    Type = DataValidationValues.List,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
                }) { Count = 1U }, worksheet.GetFirstChild<SheetData>()!);
                worksheet.Save();
            }

            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                // External links are not allowed here, and the inline list must not be mistaken for a use of Red.
                if (merge) {
                    targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                        TargetSheetNames = new Dictionary<string, string> { ["Data"] = "Copy" },
                        MaxDefinedNames = 1
                    });
                } else {
                    targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy", ExcelSheetNameValidationMode.Strict,
                        new ExcelWorksheetCopyOptions { MaxDefinedNames = 1 });
                }

                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Assert.Equal("\"Red,Green\"", GetWorksheetPartByNameForOperations(target, "Copy").Worksheet!.Descendants<Formula1>().Single().Text);
                Assert.Equal(new[] { "_xlnm.Print_Area" }, target.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Select(name => name.Name!.Value));
                Assert.Null(target.WorkbookPart.Workbook.ExternalReferences);
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData(ExcelWorksheetCopyMode.Values)]
        [InlineData(ExcelWorksheetCopyMode.Package)]
        public void Test_WorksheetCopy_PrintAreaAliasSurvivesCopyAndSave(ExcelWorksheetCopyMode mode) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"PrintAliasSource-{mode}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"PrintAliasTarget-{mode}.xlsx");
            using (var createdSource = ExcelDocument.Create(sourcePath)) {
                ExcelSheet data = createdSource.AddSheet("Data");
                for (int row = 1; row <= 3; row++) data.Cell(row, 1).SetValue(row);
                createdSource.Save();
            }

            AddDefinedName(sourcePath, "PrintRegion", "Data!$A$1:$A$3");
            AddDefinedName(sourcePath, "_xlnm.Print_Area", "PrintRegion", localSheetId: 0U);
            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = mode });
                targetDocument.Save();
            }

            using (var reloaded = ExcelDocument.Load(targetPath)) {
                // A later edit and save runs the save cleanup again.
                reloaded["Copy"].Cell(5, 1).SetValue("Later");
                reloaded.Save(new ExcelSaveOptions { SafePreflight = true });
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Workbook workbook = target.WorkbookPart!.Workbook!;
                uint copyPosition = (uint)workbook.Sheets!.Elements<Sheet>().ToList().FindIndex(sheet => sheet.Name == "Copy");
                Dictionary<string, string> names = workbook.DefinedNames!.Elements<DefinedName>()
                    .Where(name => name.LocalSheetId?.Value == copyPosition)
                    .ToDictionary(name => name.Name!.Value!, name => Unquote(name.Text));
                Assert.Equal("PrintRegion", names["_xlnm.Print_Area"]);
                Assert.Equal("Copy!$A$1:$A$3", names["PrintRegion"]);
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        private static string Unquote(string text) => text.Replace("'Copy'", "Copy");

        private static void CreateFormulaMatrixSource(string path) {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet data = document.AddSheet("My Data");
                data.Cell(1, 1).SetValue("Region");
                data.Cell(1, 2).SetValue("Amount");
                data.Cell(2, 1).SetValue("North");
                data.Cell(2, 2).SetValue(10);
                data.Cell(3, 1).SetValue("South");
                data.Cell(3, 2).SetValue(20);
                data.AddTable("A1:B3", hasHeader: true, name: "Sales");
                document.Save();
            }

            AddLinkedWorkbook(path, "Prices.xlsx");
            AddDefinedName(path, "Rate", "0.5");
            AddDefinedName(path, "Chain", "Rate*2");
            AddDefinedName(path, "Local", "'My Data'!$B$2", localSheetId: 0U);
            AddDefinedName(path, "Ext", "[1]Sheet1!$A$1");
            AddDefinedName(path, "Decoy", "[1]Sheet1!$B$1");
            AddDefinedName(path, "_xlnm.Print_Area", "OFFSET('My Data'!$A$1,0,0,ROWS(Sales[#All]),Chain)", localSheetId: 0U);
            AddDefinedName(path, "_xlnm.Print_Titles", "'My Data'!$1:$1", localSheetId: 0U);

            using SpreadsheetDocument source = SpreadsheetDocument.Open(path, true);
            Worksheet worksheet = GetWorksheetPartByNameForOperations(source, "My Data").Worksheet!;
            OpenXmlElement anchor = worksheet.GetFirstChild<SheetData>()!;
            anchor = worksheet.InsertAfter(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula(MatrixFormula)) { Type = ConditionalFormatValues.Expression, Priority = 1 },
                new ConditionalFormattingRule(new ColorScale(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Formula, Val = MatrixFormula },
                    new Color { Rgb = "FFF8696B" },
                    new Color { Rgb = "FF63BE7B" })) { Type = ConditionalFormatValues.ColorScale, Priority = 2 }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2:B3" }
            }, anchor);
            worksheet.InsertAfter(new DataValidations(new DataValidation(new Formula1(MatrixFormula)) {
                Type = DataValidationValues.Custom,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "C2:C3" }
            }) { Count = 1U }, anchor);
            worksheet.Save();
            Assert.Empty(new OpenXmlValidator().Validate(source));
        }
    }
}
