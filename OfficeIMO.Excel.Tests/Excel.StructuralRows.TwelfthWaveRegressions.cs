using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_DoesNotRewriteSpacedFunctionTokensAsCells() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.Calculation.RegisterCustomFunction(
                "ABC10",
                (_, _) => ExcelFormulaValue.FromNumber(1));
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, "LOG10 (A1)+ABC10 (A1)");

            sheet.InsertRows(5);

            Cell formulaCell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A1");
            Assert.Equal("LOG10 (A1)+ABC10 (A1)", formulaCell.CellFormula!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RemapsAndPreflightsConsolidatedPivotRangeSets() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = CreatePivotSheet(document);
                PivotTableCacheDefinitionPart cachePart = Assert.Single(
                    sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!;
                var rangeSet = new RangeSet { Sheet = "Data", Reference = "A5:B6" };
                cachePart.PivotCacheDefinition!.CacheSource = new CacheSource(
                    new Consolidation(
                        new RangeSets(rangeSet) { Count = 1U })) {
                    Type = SourceValues.Consolidation
                };
                cachePart.PivotCacheDefinition.RefreshOnLoad = false;
                cachePart.PivotCacheDefinition.SaveData = true;

                sheet.InsertRows(5);

                Assert.Equal("A6:B7", rangeSet.Reference!.Value);
                Assert.True(cachePart.PivotCacheDefinition.RefreshOnLoad!.Value);
                Assert.False(cachePart.PivotCacheDefinition.SaveData!.Value);
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = CreatePivotSheet(document);
                PivotTableCacheDefinitionPart cachePart = Assert.Single(
                    sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!;
                var rangeSet = new RangeSet {
                    Sheet = "Data",
                    Reference = $"A{A1.MaxRows}"
                };
                cachePart.PivotCacheDefinition!.CacheSource = new CacheSource(
                    new Consolidation(
                        new RangeSets(rangeSet) { Count = 1U })) {
                    Type = SourceValues.Consolidation
                };

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.InsertRows(1));

                Assert.Contains("row limit", exception.Message);
                Assert.Equal($"A{A1.MaxRows}", rangeSet.Reference!.Value);
            }
        }

        [Fact]
        public void Test_StructuralRows_PreflightsScenarioInputCells() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var input = new InputCells {
                CellReference = $"A{A1.MaxRows}",
                Val = "Keep"
            };
            sheet.WorksheetPart.Worksheet.Append(
                new Scenarios(
                    new Scenario(input) {
                        Name = "Boundary",
                        Count = 1U
                    }));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}", input.CellReference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RemapsAndPreflightsCellBackedConnectionParameters() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                Parameter parameter = AttachCellBackedConnection(document, sheet, "A5");

                sheet.InsertRows(5);

                Assert.Equal("A6", parameter.Cell!.Value);
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                Parameter parameter = AttachCellBackedConnection(
                    document,
                    sheet,
                    $"A{A1.MaxRows}");

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.InsertRows(1));

                Assert.Contains("row limit", exception.Message);
                Assert.Equal($"A{A1.MaxRows}", parameter.Cell!.Value);
            }
        }

        [Fact]
        public void Test_StructuralRows_RewritesAndPreflightsFormulaConditionalThresholds() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet data = document.AddWorksheet("Data");
                data.CellAt(5, 1).SetValue(1);
                ConditionalFormatValueObject threshold = AppendFormulaThreshold(
                    data,
                    "C5:C6",
                    "A5");

                data.InsertRows(5);

                Assert.Equal("A6", threshold.Val!.Value);
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet data = document.AddWorksheet("Data");
                ExcelSheet summary = document.AddWorksheet("Summary");
                data.CellAt(5, 1).SetValue(1);
                ConditionalFormatValueObject threshold = AppendFormulaThreshold(
                    summary,
                    "C2:C3",
                    "Data!A5");

                data.InsertRows(5);

                Assert.Equal("Data!A6", threshold.Val!.Value);
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet data = document.AddWorksheet("Data");
                ExcelSheet summary = document.AddWorksheet("Summary");
                ConditionalFormatValueObject threshold = AppendFormulaThreshold(
                    summary,
                    "C2:C3",
                    $"Data!A{A1.MaxRows}");

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => data.InsertRows(1));

                Assert.Contains("row limit", exception.Message);
                Assert.Equal($"Data!A{A1.MaxRows}", threshold.Val!.Value);
            }
        }

        private static ExcelSheet CreatePivotSheet(ExcelDocument document) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Region");
            sheet.CellAt(1, 2).SetValue("Sales");
            sheet.CellAt(2, 1).SetValue("East");
            sheet.CellAt(2, 2).SetValue(10);
            sheet.CellAt(3, 1).SetValue("West");
            sheet.CellAt(3, 2).SetValue(20);
            sheet.AddPivotTable(
                sourceRange: "A1:B3",
                destinationCell: "E10",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] {
                    new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum)
                });
            return sheet;
        }

        private static Parameter AttachCellBackedConnection(
            ExcelDocument document,
            ExcelSheet sheet,
            string cellReference) {
            var parameter = new Parameter {
                Name = "Input",
                ParameterType = ParameterValues.Cell,
                Cell = cellReference
            };
            ConnectionsPart connectionsPart = document.WorkbookPartRoot.AddNewPart<ConnectionsPart>();
            connectionsPart.Connections = new Connections(
                new Connection(
                    new Parameters(parameter) { Count = 1U }) {
                    Id = 1U,
                    Name = "Query",
                    Type = 5U,
                    RefreshedVersion = 7
                });
            QueryTablePart queryTablePart = sheet.WorksheetPart.AddNewPart<QueryTablePart>();
            queryTablePart.QueryTable = new QueryTable {
                Name = "Query",
                ConnectionId = 1U
            };
            return parameter;
        }

        private static ConditionalFormatValueObject AppendFormulaThreshold(
            ExcelSheet sheet,
            string appliesTo,
            string formula) {
            var threshold = new ConditionalFormatValueObject {
                Type = ConditionalFormatValueObjectValues.Formula,
                Val = formula
            };
            var rule = new ConditionalFormattingRule(
                new DataBar(
                    threshold,
                    new ConditionalFormatValueObject {
                        Type = ConditionalFormatValueObjectValues.Max
                    },
                    new Color { Rgb = "FF4F81BD" })) {
                Type = ConditionalFormatValues.DataBar,
                Priority = 1
            };
            sheet.WorksheetPart.Worksheet.Append(
                new ConditionalFormatting(rule) {
                    SequenceOfReferences = new ListValue<StringValue> {
                        InnerText = appliesTo
                    }
                });
            return threshold;
        }
    }
}
