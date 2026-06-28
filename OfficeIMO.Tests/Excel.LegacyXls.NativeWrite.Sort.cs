using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesSimpleSortStateMetadata() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Sort");
                    sheet.CellValue(1, 1, "Region");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(1, 3, "Date");
                    sheet.CellValue(2, 1, "West");
                    sheet.CellValue(2, 2, 20);
                    sheet.CellValue(2, 3, "2026-01-02");

                    Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                    worksheet.RemoveAllChildren<SortState>();
                    worksheet.AppendChild(new SortState(
                        new SortCondition { Reference = "A2:A4", Descending = true },
                        new SortCondition { Reference = "B2:B4" },
                        new SortCondition { Reference = "C2:C4", Descending = true }) {
                        Reference = "A1:C4",
                        ColumnSort = true,
                        CaseSensitive = true,
                        SortMethod = SortMethodValues.PinYin
                    });
                    worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheetResult = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsSortSettings sortSettings = Assert.IsType<LegacyXlsSortSettings>(worksheetResult.SortSettings);
                Assert.True(sortSettings.SortLeftToRight);
                Assert.True(sortSettings.Key1Descending);
                Assert.False(sortSettings.Key2Descending);
                Assert.True(sortSettings.Key3Descending);
                Assert.True(sortSettings.CaseSensitive);
                Assert.Equal(0, sortSettings.CustomListIndex);
                Assert.True(sortSettings.UsePhoneticInformation);
                Assert.Equal("A2:A4", sortSettings.Key1);
                Assert.Equal("B2:B4", sortSettings.Key2);
                Assert.Equal("C2:C4", sortSettings.Key3);
                Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.Sort]);

                SortState projectedSort = Assert.IsType<SortState>(result.Document.Sheets[0].WorksheetPart.Worksheet!.GetFirstChild<SortState>());
                Assert.Equal("A2:C4", projectedSort.Reference!.Value);
                Assert.True(projectedSort.ColumnSort!.Value);
                Assert.True(projectedSort.CaseSensitive!.Value);
                Assert.Equal(SortMethodValues.PinYin, projectedSort.SortMethod!.Value);
                SortCondition[] projectedConditions = projectedSort.Elements<SortCondition>().ToArray();
                Assert.Equal(3, projectedConditions.Length);
                Assert.Equal("A2:A4", projectedConditions[0].Reference!.Value);
                Assert.True(projectedConditions[0].Descending!.Value);
                Assert.Equal("B2:B4", projectedConditions[1].Reference!.Value);
                Assert.Null(projectedConditions[1].Descending);
                Assert.Equal("C2:C4", projectedConditions[2].Reference!.Value);
                Assert.True(projectedConditions[2].Descending!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSortStatesWithMoreThanThreeKeys() {
            AssertNativeXlsSaveNotSupported("sort states with more than three sort keys", (document, sheet) => {
                sheet.CellValue(1, 1, "Sort");
                Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                worksheet.AppendChild(new SortState(
                    new SortCondition { Reference = "A2:A4" },
                    new SortCondition { Reference = "B2:B4" },
                    new SortCondition { Reference = "C2:C4" },
                    new SortCondition { Reference = "D2:D4" }) {
                    Reference = "A1:D4"
                });
                worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksColorIconAndCustomListSortStates() {
            AssertNativeXlsSaveNotSupported("sort states with color or icon sort conditions", (document, sheet) => {
                sheet.CellValue(1, 1, "Sort");
                Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                worksheet.AppendChild(new SortState(
                    new SortCondition {
                        Reference = "A2:A4",
                        SortBy = SortByValues.CellColor,
                        FormatId = 1U
                    }) {
                    Reference = "A1:A4"
                });
                worksheet.Save();
            });

            AssertNativeXlsSaveNotSupported("sort states with custom-list sort conditions", (document, sheet) => {
                sheet.CellValue(1, 1, "Sort");
                Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                worksheet.AppendChild(new SortState(
                    new SortCondition {
                        Reference = "A2:A4",
                        CustomList = "High,Medium,Low"
                    }) {
                    Reference = "A1:A4"
                });
                worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSortStateUnsupportedMetadata() {
            AssertNativeXlsSaveNotSupported("sort states with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Sort");
                Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                var sortState = new SortState(
                    new SortCondition { Reference = "A2:A4" }) {
                    Reference = "A1:A4"
                };
                sortState.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                worksheet.AppendChild(sortState);
                worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSortConditionUnsupportedMetadata() {
            AssertNativeXlsSaveNotSupported("sort states with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Sort");
                Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                var condition = new SortCondition { Reference = "A2:A4" };
                condition.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                worksheet.AppendChild(new SortState(condition) {
                    Reference = "A1:A4"
                });
                worksheet.Save();
            });
        }
    }
}
