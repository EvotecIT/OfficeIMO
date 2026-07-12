using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesDefaultTableStyleNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorkSheet("TableStyles");
                    sheet.CellValue(1, 1, "Default table style names");

                    WorkbookStylesPart stylesPart = document.WorkbookPartRoot.WorkbookStylesPart
                        ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
                    Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                    TableStyles tableStyles = stylesheet.TableStyles ??= new TableStyles();
                    tableStyles.DefaultTableStyle = "TableStyleDark3";
                    tableStyles.DefaultPivotStyle = "PivotStyleMedium4";
                    tableStyles.Count = 0U;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsTableStyleCollection collection = Assert.Single(result.Workbook.TableStyleCollections);
                Assert.Equal("TableStyleDark3", collection.DefaultTableStyleName);
                Assert.Equal("PivotStyleMedium4", collection.DefaultPivotStyleName);

                TableStyles projectedTableStyles = result.Document.WorkbookPartRoot
                    .WorkbookStylesPart!
                    .Stylesheet!
                    .TableStyles!;
                Assert.Equal("TableStyleDark3", projectedTableStyles.DefaultTableStyle!.Value);
                Assert.Equal("PivotStyleMedium4", projectedTableStyles.DefaultPivotStyle!.Value);
                Assert.Equal(0U, projectedTableStyles.Count!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCustomWorkbookTableStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorkSheet("TableStyles");
                    sheet.CellValue(1, 1, "Custom workbook table style");

                    WorkbookStylesPart stylesPart = document.WorkbookPartRoot.WorkbookStylesPart
                        ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
                    Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                    stylesheet.DifferentialFormats = new DifferentialFormats(
                        new DifferentialFormat(),
                        new DifferentialFormat(),
                        new DifferentialFormat(),
                        new DifferentialFormat(),
                        new DifferentialFormat()) {
                        Count = 5U
                    };

                    TableStyles tableStyles = stylesheet.TableStyles ??= new TableStyles();
                    tableStyles.DefaultTableStyle = "TableStyleMedium2";
                    tableStyles.DefaultPivotStyle = "PivotStyleLight16";
                    tableStyles.RemoveAllChildren<DocumentFormat.OpenXml.Spreadsheet.TableStyle>();

                    var customStyle = new DocumentFormat.OpenXml.Spreadsheet.TableStyle {
                        Name = "OfficeIMOCustomTableStyle",
                        Table = true,
                        Pivot = false,
                        Count = 2U
                    };
                    customStyle.Append(
                        new TableStyleElement {
                            Type = TableStyleValues.HeaderRow,
                            FormatId = 3U
                        },
                        new TableStyleElement {
                            Type = TableStyleValues.FirstRowStripe,
                            Size = 2U,
                            FormatId = 4U
                        });
                    tableStyles.Append(customStyle);
                    tableStyles.Count = 1U;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();

                LegacyXlsTableStyleCollection collection = Assert.Single(result.Workbook.TableStyleCollections);
                Assert.Equal(146U, collection.TotalStyleCount);
                Assert.Equal("TableStyleMedium2", collection.DefaultTableStyleName);
                Assert.Equal("PivotStyleLight16", collection.DefaultPivotStyleName);

                LegacyXlsTableStyle style = Assert.Single(result.Workbook.TableStyles);
                Assert.Equal("OfficeIMOCustomTableStyle", style.Name);
                Assert.True(style.AppliesToTables);
                Assert.False(style.AppliesToPivotTables);
                Assert.Equal(2U, style.DeclaredElementCount);

                Assert.Collection(
                    style.Elements,
                    element => {
                        Assert.Equal("HeaderRow", element.ElementTypeName);
                        Assert.Equal(0U, element.StripeSize);
                        Assert.Equal(3U, element.DifferentialFormatIndex);
                    },
                    element => {
                        Assert.Equal("RowStripe1", element.ElementTypeName);
                        Assert.Equal(2U, element.StripeSize);
                        Assert.Equal(4U, element.DifferentialFormatIndex);
                    });

                Assert.Equal(1, result.ImportReport.TableStyleDefinitionCount);
                Assert.Equal(2, result.ImportReport.TableStyleElementRecordCount);
                Assert.Equal(1, result.ImportReport.TableStylesByName["OfficeIMOCustomTableStyle"]);
                Assert.Equal(1, result.ImportReport.TableStyleElementsByType["HeaderRow"]);
                Assert.Equal(1, result.ImportReport.TableStyleElementsByType["RowStripe1"]);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
