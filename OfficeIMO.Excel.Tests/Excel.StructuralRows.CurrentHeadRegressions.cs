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
