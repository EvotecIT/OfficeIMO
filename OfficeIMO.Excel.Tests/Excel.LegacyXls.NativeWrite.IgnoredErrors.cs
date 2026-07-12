using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesIgnoredErrors() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorkSheet("IgnoredErrors");
                    sheet.CellValue(1, 1, "1");
                    sheet.CellValue(2, 1, "2");
                    sheet.CellValue(3, 3, "3");

                    sheet.WorksheetPart.Worksheet.Append(new IgnoredErrors(
                        new IgnoredError {
                            SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:B2" },
                            EvalError = true,
                            EmptyCellReference = true,
                            NumberStoredAsText = true,
                            FormulaRange = true,
                            Formula = true,
                            TwoDigitTextYear = true,
                            UnlockedFormula = true,
                            ListDataValidation = true
                        },
                        new IgnoredError {
                            SequenceOfReferences = new ListValue<StringValue> { InnerText = "C3" },
                            NumberStoredAsText = true
                        }));
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal(2, worksheet.IgnoredErrors.Count);
                LegacyXlsIgnoredError allFlags = worksheet.IgnoredErrors[0];
                Assert.Equal(new[] { "A1:B2" }, allFlags.References);
                Assert.True(allFlags.EvaluationError);
                Assert.True(allFlags.EmptyCellReference);
                Assert.True(allFlags.NumberStoredAsText);
                Assert.True(allFlags.FormulaRange);
                Assert.True(allFlags.Formula);
                Assert.True(allFlags.TwoDigitTextYear);
                Assert.True(allFlags.UnlockedFormula);
                Assert.True(allFlags.ListDataValidation);

                LegacyXlsIgnoredError numberAsText = worksheet.IgnoredErrors[1];
                Assert.Equal(new[] { "C3" }, numberAsText.References);
                Assert.True(numberAsText.NumberStoredAsText);
                Assert.False(numberAsText.EvaluationError);

                IgnoredError[] projectedErrors = result.Document.Sheets.Single()
                    .WorksheetPart.Worksheet
                    .Elements<IgnoredErrors>()
                    .Single()
                    .Elements<IgnoredError>()
                    .ToArray();
                Assert.Equal(2, projectedErrors.Length);
                Assert.Equal("A1:B2", projectedErrors[0].SequenceOfReferences!.InnerText);
                Assert.True(projectedErrors[0].EvalError!.Value);
                Assert.True(projectedErrors[0].EmptyCellReference!.Value);
                Assert.True(projectedErrors[0].NumberStoredAsText!.Value);
                Assert.True(projectedErrors[0].FormulaRange!.Value);
                Assert.True(projectedErrors[0].Formula!.Value);
                Assert.True(projectedErrors[0].TwoDigitTextYear!.Value);
                Assert.True(projectedErrors[0].UnlockedFormula!.Value);
                Assert.True(projectedErrors[0].ListDataValidation!.Value);
                Assert.Equal("C3", projectedErrors[1].SequenceOfReferences!.InnerText);
                Assert.True(projectedErrors[1].NumberStoredAsText!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
