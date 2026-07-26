using System;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_FormulaQuoteScanningRemainsLinear() {
            string quotedSheetName = "'" + new string('"', 50_000) + "'";
            string formula = quotedSheetName + "!A5+\"A5\"+A5";
            var stopwatch = Stopwatch.StartNew();

            string rewritten = ExcelSheet.RewriteFormulaReferencesOutsideStrings(
                formula,
                segment => segment.Replace("A5", "A6"));
            stopwatch.Stop();

            Assert.Equal(quotedSheetName + "!A6+\"A5\"+A6", rewritten);
            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(5),
                $"Formula scanning took {stopwatch.Elapsed}.");
        }

        [Fact]
        public void Test_StructuralRows_IgnoresMalformedImplicitCellsWithoutSharedFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet edited = document.AddWorksheet("Edited");
            edited.CellAt(1, 1).SetValue("Keep");
            ExcelSheet malformed = document.AddWorksheet("Malformed");
            malformed.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Append(
                new Row(new Cell { CellValue = new CellValue("value") }) {
                    RowIndex = uint.MaxValue
                });

            edited.InsertRows(1);

            Assert.Equal("Keep", edited.CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_HandlesMalformedImplicitSharedFormulaCoordinatesWithoutOverflow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet edited = document.AddWorksheet("Edited");
            ExcelSheet malformed = document.AddWorksheet("Malformed");
            malformed.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Append(
                new Row(new Cell {
                    CellFormula = new CellFormula("A1") {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 9U,
                        Reference = "A1:A1"
                    }
                }) {
                    RowIndex = uint.MaxValue
                });

            edited.InsertRows(1);

            CellFormula formula = Assert.Single(
                malformed.WorksheetPart.Worksheet.Descendants<CellFormula>());
            Assert.Equal("A1", formula.Text);
        }

        [Fact]
        public void Test_StructuralRows_IgnoresSameNamedDefinitionsFromOtherSheetScopes() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = CreatePivotSheet(document);
            document.AddWorksheet("Other");
            WorksheetSource source = Assert.Single(
                data.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!
                .PivotCacheDefinition!.CacheSource!.WorksheetSource!;
            source.Reference = null;
            source.Sheet = null;
            source.Name = "PivotSource";
            var globalName = new DefinedName("'Data'!$A$5:$B$7") {
                Name = "PivotSource"
            };
            var otherLocalName = new DefinedName("$A$2:$B$3") {
                Name = "PivotSource",
                LocalSheetId = 1U
            };
            document.WorkbookRoot.DefinedNames = new DefinedNames(
                globalName,
                otherLocalName);

            data.DeleteRows(2);

            Assert.Equal("'Data'!$A$4:$B$6", globalName.Text);
            Assert.Equal("$A$2:$B$3", otherLocalName.Text);
        }

        [Fact]
        public void Test_StructuralRows_PreservesActiveCellIdForSingleCellSelections() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var selection = new Selection {
                ActiveCell = "B6",
                ActiveCellId = 1U,
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = "A3 B6"
                }
            };
            var views = new SheetViews(
                new SheetView(selection) { WorkbookViewId = 0U });
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheet.WorksheetPart.Worksheet.InsertBefore(views, sheetData);

            sheet.InsertRows(4);

            Assert.Equal("B7", selection.ActiveCell!.Value);
            Assert.Equal("A3 B7", selection.SequenceOfReferences!.InnerText);
            Assert.Equal(1U, selection.ActiveCellId!.Value);
        }
    }
}
