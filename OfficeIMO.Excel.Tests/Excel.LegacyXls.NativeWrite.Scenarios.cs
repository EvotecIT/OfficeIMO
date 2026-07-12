using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetScenarios() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorkSheet("Scenarios");
                    sheet.CellValue(1, 1, "Base");
                    sheet.CellValue(2, 1, "Other");
                    sheet.WorksheetPart.Worksheet.Append(new Scenarios(
                        new Scenario(
                            new InputCells { CellReference = "A1", Val = "Optimistic" },
                            new InputCells { CellReference = "A2", Val = "12", Deleted = true }) {
                            Name = "Best case",
                            User = "OfficeIMO",
                            Comment = "Roundtrip scenario",
                            Locked = true,
                            Count = 2U
                        }) {
                        Current = 0U,
                        Show = 0U,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A2" }
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                Assert.NotNull(worksheet.ScenarioManager);
                Assert.Equal(1, worksheet.ScenarioManager!.ScenarioCount);
                Assert.Equal(0, worksheet.ScenarioManager.CurrentScenarioIndex);
                Assert.Equal(0, worksheet.ScenarioManager.ShownScenarioIndex);
                Assert.Equal(new[] { "A1:A2" }, worksheet.ScenarioManager.ResultRanges);

                LegacyXlsScenario scenario = Assert.Single(worksheet.Scenarios);
                Assert.Equal("Best case", scenario.Name);
                Assert.True(scenario.Locked);
                Assert.False(scenario.Hidden);
                Assert.Equal("OfficeIMO", scenario.User);
                Assert.Equal("Roundtrip scenario", scenario.Comment);
                Assert.Equal(2, scenario.InputCells.Count);
                Assert.Equal("A1", scenario.InputCells[0].CellReference);
                Assert.Equal("Optimistic", scenario.InputCells[0].Value);
                Assert.False(scenario.InputCells[0].Deleted);
                Assert.Equal("A2", scenario.InputCells[1].CellReference);
                Assert.Equal("12", scenario.InputCells[1].Value);
                Assert.True(scenario.InputCells[1].Deleted);

                Scenarios projectedScenarios = result.Document.Sheets.Single()
                    .WorksheetPart.Worksheet
                    .Elements<Scenarios>()
                    .Single();
                Assert.Equal("A1:A2", projectedScenarios.SequenceOfReferences!.InnerText);
                Scenario projectedScenario = Assert.Single(projectedScenarios.Elements<Scenario>());
                Assert.Equal("Best case", projectedScenario.Name!.Value);
                Assert.Equal("OfficeIMO", projectedScenario.User!.Value);
                Assert.Equal("Roundtrip scenario", projectedScenario.Comment!.Value);
                InputCells[] projectedInputCells = projectedScenario.Elements<InputCells>().ToArray();
                Assert.Equal("A1", projectedInputCells[0].CellReference!.Value);
                Assert.Equal("Optimistic", projectedInputCells[0].Val!.Value);
                Assert.Equal("A2", projectedInputCells[1].CellReference!.Value);
                Assert.True(projectedInputCells[1].Deleted!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedWorksheetScenarioPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet scenario payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Scenario value");
                var scenario = new Scenario {
                    Name = "Large scenario",
                    Count = 32U
                };

                string largeValue = new string('A', 2000);
                for (int row = 1; row <= 32; row++) {
                    scenario.Append(new InputCells {
                        CellReference = "A" + row,
                        Val = largeValue
                    });
                }

                sheet.WorksheetPart.Worksheet.Append(new Scenarios(scenario));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedWorksheetScenarioMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet scenarios with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Scenario value");
                var scenario = new Scenario(
                    new InputCells {
                        CellReference = "A1",
                        Val = "Optimistic"
                    }) {
                    Name = "Best case",
                    Count = 1U
                };
                scenario.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));

                sheet.WorksheetPart.Worksheet.Append(new Scenarios(scenario));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }
    }
}
