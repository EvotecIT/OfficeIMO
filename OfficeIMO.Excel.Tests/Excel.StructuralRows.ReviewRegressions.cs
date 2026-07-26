using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RewritesPostfixReferencesWithoutTreatingDefinedNamesAsCells() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(3, 1).SetValue(10);
            sheet.CellFormula(1, 2, "A3%+A3#+XFE1+SUM(XFE1:XFE2)");

            sheet.InsertRows(3);

            Assert.Equal("A4%+A4#+XFE1+SUM(XFE1:XFE2)", sheet.GetFormulaText(1, 2));
        }

        [Fact]
        public void Test_StructuralRows_RebasesAnchoredValidationAndConditionalFormattingFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.AnchoredFormulas.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellAt(2, 1).SetValue(1);
                sheet.CellAt(3, 1).SetValue(2);
                sheet.Range("B2:B3").Validation.CustomFormula("A2>$A$2");
                sheet.AddConditionalFormulaRule("C2:C3", "A2>$A$2");

                sheet.InsertRows(2);

                DataValidation validation = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<DataValidation>());
                ConditionalFormatting conditional = Assert.Single(
                    sheet.WorksheetPart.Worksheet.Elements<ConditionalFormatting>());
                Assert.Equal("B3:B4", validation.SequenceOfReferences!.InnerText);
                Assert.Equal("A3>$A$3", validation.Formula1!.Text);
                Assert.Equal("C3:C4", conditional.SequenceOfReferences!.InnerText);
                Assert.Equal("A3>$A$3", Assert.Single(conditional.Descendants<Formula>()).Text);

                sheet.DeleteRows(3);
                Assert.Equal("B3", validation.SequenceOfReferences!.InnerText);
                Assert.Equal("A3>#REF!", validation.Formula1!.Text);
                Assert.Equal("C3", conditional.SequenceOfReferences!.InnerText);
                Assert.Equal("A3>#REF!", Assert.Single(conditional.Descendants<Formula>()).Text);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void Test_StructuralRows_RemapsProtectedSortAndDataTableReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.RangeMetadata.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellAt(2, 1).SetValue(1);
                sheet.CellAt(3, 1).SetValue(2);
                sheet.CellAt(4, 2).SetValue(3);
                sheet.CellAt(1, 4).SetValue("Key");
                sheet.CellAt(1, 5).SetValue("Value");
                sheet.CellAt(2, 4).SetValue("A");
                sheet.CellAt(2, 5).SetValue(1);
                sheet.CellAt(3, 4).SetValue("B");
                sheet.CellAt(3, 5).SetValue(2);
                sheet.CellAt(4, 4).SetValue("C");
                sheet.CellAt(4, 5).SetValue(3);
                sheet.Protect();

                var protectedRange = new ProtectedRange {
                    Name = "Editable",
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A3" }
                };
                SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                sheet.WorksheetPart.Worksheet.InsertAfter(
                    new ProtectedRanges(protectedRange),
                    sheetData);

                sheet.AutoFilterAdd("A1:B4");
                AutoFilter filter = sheet.WorksheetPart.Worksheet.GetFirstChild<AutoFilter>()!;
                filter.Append(new SortState(
                    new SortCondition { Reference = "A2:A3" }) {
                    Reference = "A2:B3"
                });
                sheet.AddTable(
                    "D1:E4",
                    hasHeader: true,
                    name: "SortedData",
                    OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table;
                table.GetFirstChild<AutoFilter>()!.Append(new SortState(
                    new SortCondition { Reference = "D2:D4" }) {
                    Reference = "D2:E4"
                });

                Row row = sheetData.Elements<Row>().Single(item => item.RowIndex?.Value == 2U);
                row.Append(new Cell {
                    CellReference = "C2",
                    CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.DataTable,
                        Reference = "C2:C3",
                        R1 = "A3",
                        R2 = "B4"
                    }
                });

                sheet.InsertRows(2);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart part = GetStructuralWorksheetPart(spreadsheet, "Data");
                ProtectedRange protectedRange = Assert.Single(part.Worksheet.Descendants<ProtectedRange>());
                SortState sortState = Assert.Single(part.Worksheet.Descendants<SortState>());
                SortCondition condition = Assert.Single(sortState.Elements<SortCondition>());
                SortState tableSortState = Assert.Single(
                    Assert.Single(part.TableDefinitionParts).Table.Descendants<SortState>());
                SortCondition tableCondition = Assert.Single(tableSortState.Elements<SortCondition>());
                CellFormula dataTable = part.Worksheet.Descendants<CellFormula>()
                    .Single(formula => formula.FormulaType?.Value == CellFormulaValues.DataTable);

                Assert.Equal("A3:A4", protectedRange.SequenceOfReferences!.InnerText);
                Assert.Equal("A3:B4", sortState.Reference!.Value);
                Assert.Equal("A3:A4", condition.Reference!.Value);
                Assert.Equal("D3:E5", tableSortState.Reference!.Value);
                Assert.Equal("D3:D5", tableCondition.Reference!.Value);
                Assert.Equal("C3:C4", dataTable.Reference!.Value);
                Assert.Equal("A4", dataTable.R1!.Value);
                Assert.Equal("B5", dataTable.R2!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_StructuralRows_PreservesExternalLinkLocationsAndHonorsOwnedRowMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Header");
            sheet.CellAt(2, 1).SetValue("Value");
            sheet.CellAt(3, 1).SetValue("Total");
            sheet.SetHyperlink(2, 1, "https://example.org", display: "External");
            Hyperlink hyperlink = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Hyperlink>());
            hyperlink.Location = "Data!A2";

            sheet.AddTable(
                "A1:A3",
                hasHeader: true,
                name: "TotalsData",
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table;
            table.TotalsRowCount = 1U;
            table.TotalsRowShown = false;

            Assert.Throws<InvalidOperationException>(() => sheet.DeleteRows(3));

            table.TotalsRowCount = 0U;
            SheetDimension dimension = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetDimension>()
                ?? sheet.WorksheetPart.Worksheet.InsertAt(new SheetDimension(), 0);
            dimension.Reference = $"A1:A{A1.MaxRows}";
            sheet.InsertRows(2);

            Assert.Equal("A3", hyperlink.Reference!.Value);
            Assert.Equal("Data!A2", hyperlink.Location!.Value);
            Assert.Equal("External", sheet.CellAt(3, 1).GetValue<string>());
        }
    }
}
