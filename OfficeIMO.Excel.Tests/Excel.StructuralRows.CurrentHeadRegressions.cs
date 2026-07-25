using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_PreservesOutOfGridNamesStructuredReferencesAndExternalTargets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("A3 Total");
            sheet.CellAt(2, 1).SetValue(10);
            sheet.AddTable(
                "A1:A2",
                hasHeader: true,
                name: "Table1",
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            sheet.CellFormula(1, 3, "A1048577+SUM(Table1[A3 Total])");
            sheet.SetHyperlink(2, 1, "https://example.org/book.xlsx", display: "External");
            Hyperlink hyperlink = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Hyperlink>());
            hyperlink.Location = $"Data!A{A1.MaxRows}";

            sheet.InsertRows(2);

            Assert.Equal("A1048577+SUM(Table1[A3 Total])", sheet.GetFormulaText(1, 3));
            Assert.Equal("A3", hyperlink.Reference!.Value);
            Assert.Equal($"Data!A{A1.MaxRows}", hyperlink.Location!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RemapsOffice2010DataValidationTargetsAndFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.CellAt(3, 1).SetValue(2);

            var validation = new X14.DataValidation(
                new X14.DataValidationForumla1(new Xm.Formula("A2>0")),
                new Xm.ReferenceSequence("B2:B3"));
            var validations = new X14.DataValidations(validation) { Count = 1U };
            sheet.WorksheetPart.Worksheet.Append(
                new ExtensionList(
                    new Extension(validations) {
                        Uri = "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"
                    }));

            sheet.InsertRows(2);

            Assert.Equal("B3:B4", validation.ReferenceSequence!.Text);
            Assert.Equal("A3>0", validation.DataValidationForumla1!.Formula!.Text);
            Assert.Equal(1U, validations.Count!.Value);
        }

        [Fact]
        public void Test_StructuralRows_DoesNotRebaseRelativeReferencesToOtherSheets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            sheet.CellAt(2, 1).SetValue(1);

            var validation = new DataValidation(
                new Formula1("Other!A2+SUM(Other!B2:B3)+SUM(Other!2:3)>0"),
                new Formula2("A2>0")) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2" }
            };
            sheet.WorksheetPart.Worksheet.Append(
                new DataValidations(validation) { Count = 1U });

            sheet.InsertRows(2);

            Assert.Equal("B3", validation.SequenceOfReferences!.InnerText);
            Assert.Equal("Other!A2+SUM(Other!B2:B3)+SUM(Other!2:3)>0", validation.Formula1!.Text);
            Assert.Equal("A3>0", validation.Formula2!.Text);

            sheet.DeleteRows(2);

            Assert.Equal("B2", validation.SequenceOfReferences!.InnerText);
            Assert.Equal("Other!A2+SUM(Other!B2:B3)+SUM(Other!2:3)>0", validation.Formula1!.Text);
            Assert.Equal("A2>0", validation.Formula2!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RemapsStandardAndOffice2010IgnoredErrors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue("1");
            sheet.CellAt(6, 1).SetValue("2");

            var standardError = new IgnoredError {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A5:A6" },
                NumberStoredAsText = true
            };
            sheet.WorksheetPart.Worksheet.Append(new IgnoredErrors(standardError));

            var extendedError = new X14.IgnoredError(
                new Xm.ReferenceSequence("B5:B6")) {
                NumberStoredAsText = true
            };
            sheet.WorksheetPart.Worksheet.Append(
                new ExtensionList(
                    new Extension(new X14.IgnoredErrors(extendedError)) {
                        Uri = "{01252117-D84E-4E92-8308-4BE1C098FCBB}"
                    }));

            sheet.InsertRows(5);

            Assert.Equal("A6:A7", standardError.SequenceOfReferences!.InnerText);
            Assert.Equal("B6:B7", extendedError.ReferenceSequence!.Text);

            sheet.DeleteRows(5);

            Assert.Equal("A5:A6", standardError.SequenceOfReferences!.InnerText);
            Assert.Equal("B5:B6", extendedError.ReferenceSequence!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RemapsOffice2010ConditionalFormattingTargetsAndFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            sheet.CellAt(2, 1).SetValue(1);

            var rule = new X14.ConditionalFormattingRule(
                new Xm.Formula("A2>0"),
                new Xm.Formula("Other!A2>0")) {
                Type = ConditionalFormatValues.Expression,
                Priority = 1,
                Id = "{83B1F34C-CC5E-4C1E-BE8B-2E25B97CDE54}"
            };
            var formatting = new X14.ConditionalFormatting(
                rule,
                new Xm.ReferenceSequence("C2:C3"));
            var formattings = new X14.ConditionalFormattings(formatting);
            sheet.WorksheetPart.Worksheet.Append(
                new ExtensionList(
                    new Extension(formattings) {
                        Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                    }));

            sheet.InsertRows(2);

            Assert.Equal("C3:C4", formatting.GetFirstChild<Xm.ReferenceSequence>()!.Text);
            Xm.Formula[] formulas = rule.Elements<Xm.Formula>().ToArray();
            Assert.Equal("A3>0", formulas[0].Text);
            Assert.Equal("Other!A2>0", formulas[1].Text);

            sheet.DeleteRows(2);

            Assert.Equal("C2:C3", formatting.GetFirstChild<Xm.ReferenceSequence>()!.Text);
            Assert.Equal("A2>0", formulas[0].Text);
            Assert.Equal("Other!A2>0", formulas[1].Text);
        }

        [Fact]
        public void Test_StructuralRows_RejectsFormControlsBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue("Keep");
            sheet.WorksheetPart.Worksheet.Append(new Controls());

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(2));

            Assert.Contains("form controls", exception.Message, System.StringComparison.OrdinalIgnoreCase);
            Assert.Equal("Keep", sheet.CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void Test_DeleteRows_ValidatesAllSharedFormulaGroupsBeforeMaterializingAny() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet first = document.AddWorksheet("First");
            ExcelSheet second = document.AddWorksheet("Second");
            first.CellAt(2, 1).SetValue(1);
            first.CellAt(3, 1).SetValue(2);
            second.CellAt(2, 1).SetValue(1);
            second.CellAt(3, 1).SetValue(2);

            AppendSharedFormulaGroup(first, sharedIndex: 31U, anchorReference: "B2:B3");
            AppendSharedFormulaGroup(second, sharedIndex: 32U, anchorReference: "B2:B2");

            Assert.Throws<InvalidOperationException>(() => first.DeleteRows(1));

            CellFormula[] firstFormulas = first.WorksheetPart.Worksheet.Descendants<CellFormula>().ToArray();
            Assert.All(firstFormulas, formula => Assert.Equal(CellFormulaValues.Shared, formula.FormulaType!.Value));
            Assert.Equal(31U, firstFormulas[0].SharedIndex!.Value);
            Assert.Equal(31U, firstFormulas[1].SharedIndex!.Value);
            Assert.Equal("A2*2", firstFormulas[0].Text);
            Assert.True(string.IsNullOrEmpty(firstFormulas[1].Text));
        }

        [Fact]
        public void Test_StructuralRows_InvalidatesNameBasedPivotSources() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Region");
            sheet.CellAt(1, 2).SetValue("Sales");
            sheet.CellAt(2, 1).SetValue("East");
            sheet.CellAt(2, 2).SetValue(10);
            sheet.CellAt(3, 1).SetValue("West");
            sheet.CellAt(3, 2).SetValue(20);
            sheet.AddTable(
                "A1:B3",
                hasHeader: true,
                name: "PivotSourceData",
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            sheet.AddPivotTable(
                sourceRange: "A1:B3",
                destinationCell: "E10",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

            PivotTablePart pivotPart = Assert.Single(sheet.WorksheetPart.PivotTableParts);
            PivotTableCacheDefinitionPart cachePart = pivotPart.PivotTableCacheDefinitionPart!;
            WorksheetSource source = cachePart.PivotCacheDefinition!.CacheSource!.WorksheetSource!;
            source.Name = "PivotSourceData";
            source.Reference = null;
            source.Sheet = null;
            cachePart.PivotCacheDefinition.RefreshOnLoad = false;
            cachePart.PivotCacheDefinition.SaveData = true;

            sheet.InsertRows(3);

            Assert.True(cachePart.PivotCacheDefinition.RefreshOnLoad!.Value);
            Assert.False(cachePart.PivotCacheDefinition.SaveData!.Value);
            Assert.Equal(
                0U,
                cachePart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Count!.Value);
        }

        private static void AppendSharedFormulaGroup(
            ExcelSheet sheet,
            uint sharedIndex,
            string anchorReference) {
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            Row firstRow = sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 2U);
            Row secondRow = sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 3U);
            firstRow.Append(new Cell {
                CellReference = "B2",
                CellFormula = new CellFormula("A2*2") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = sharedIndex,
                    Reference = anchorReference
                }
            });
            secondRow.Append(new Cell {
                CellReference = "B3",
                CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = sharedIndex
                }
            });
        }
    }
}
