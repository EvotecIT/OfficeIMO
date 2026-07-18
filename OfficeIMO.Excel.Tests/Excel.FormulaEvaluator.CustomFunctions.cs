using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void FormulaEvaluator_CalculatesRegisteredCustomFunctions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelCustomFormulaFunctionContext? labelContext = null;

            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.Calculation.RegisterCustomFunction("GROSSMARGIN", (_, arguments) => {
                        if (arguments.Count != 2
                            || arguments[0].Kind != ExcelFormulaValueKind.Number
                            || arguments[1].Kind != ExcelFormulaValueKind.Number
                            || arguments[0].Number == 0d) {
                            return null;
                        }

                        return ExcelFormulaValue.FromNumber((arguments[0].Number - arguments[1].Number) / arguments[0].Number);
                    });
                    document.Calculation.RegisterCustomFunction("RANGESUM", (_, arguments) =>
                        arguments.All(argument => argument.Kind == ExcelFormulaValueKind.Number)
                            ? ExcelFormulaValue.FromNumber(arguments.Sum(argument => argument.Number))
                            : null);
                    document.Calculation.RegisterCustomFunction("LABEL", (context, arguments) => {
                        labelContext = context;
                        if (arguments.Count != 2
                            || arguments[0].Kind != ExcelFormulaValueKind.Text
                            || arguments[1].Kind != ExcelFormulaValueKind.Number) {
                            return null;
                        }

                        return ExcelFormulaValue.FromText(arguments[0].Text + ":" + arguments[1].Number.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    });
                    document.Calculation.RegisterCustomFunction("CUSTOMERROR", (_, arguments) =>
                        arguments.Count == 0 ? ExcelFormulaValue.FromError("#N/A") : null);

                    ExcelSheet sheet = document.AddWorksheet("Report");
                    sheet.CellValue(1, 1, 120d);
                    sheet.CellValue(2, 1, 90d);
                    sheet.CellValue(3, 1, 30d);
                    sheet.CellFormula(1, 2, "GROSSMARGIN(A1,A2)");
                    sheet.CellFormula(2, 2, "RANGESUM(A1:A3)");
                    sheet.CellFormula(3, 2, "LABEL(\"Margin\",B1)");
                    sheet.CellFormula(4, 2, "CUSTOMERROR()");

                    ExcelFormulaInspection before = sheet.InspectFormulas();
                    Assert.Equal(4, before.TotalFormulas);
                    Assert.Equal(4, before.SupportedFormulas);
                    Assert.Contains(before.Formulas, formula =>
                        formula.CellReference == "B3"
                        && formula.IsSupportedByOfficeIMO
                        && formula.HasDependencyIssues);
                    Assert.Equal(new[] { "CUSTOMERROR", "GROSSMARGIN", "LABEL", "RANGESUM" }, document.Calculation.CustomFunctionNames);

                    Assert.Equal(4, document.Calculate());
                    ExcelFormulaInspection after = sheet.InspectFormulas();
                    Assert.Equal(4, after.SupportedFormulas);
                    Assert.Contains(after.Formulas, formula => formula.CellReference == "B1" && formula.CachedValue == "0.25");
                    Assert.Contains(after.Formulas, formula => formula.CellReference == "B2" && formula.CachedValue == "240");
                    Assert.Contains(after.Formulas, formula => formula.CellReference == "B3" && formula.CachedValue == "Margin:0.25");
                    Assert.Contains(after.Formulas, formula => formula.CellReference == "B4" && formula.CachedValue == "#N/A");
                    Assert.NotNull(labelContext);
                    Assert.Same(document, labelContext!.Workbook);
                    Assert.Equal("Report", labelContext.Worksheet.Name);
                    Assert.Equal("LABEL", labelContext.FunctionName);
                    Assert.Equal("B3", labelContext.CellReference);

                    Assert.True(document.Calculation.RemoveCustomFunction("grossmargin"));
                    ExcelFormulaCellInfo unsupported = Assert.Single(sheet.InspectFormulas().Formulas, formula => formula.CellReference == "B1");
                    Assert.False(unsupported.IsSupportedByOfficeIMO);
                    Assert.Contains("GROSSMARGIN", unsupported.UnsupportedReason);
                    document.Save();
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
                Cell[] formulas = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellFormula != null)
                    .ToArray();
                Assert.Equal(4, formulas.Length);
                Assert.Contains(formulas, cell => cell.CellReference?.Value == "B3" && cell.CellValue?.InnerText == "Margin:0.25");
                Assert.Contains(formulas, cell => cell.CellReference?.Value == "B4" && cell.DataType?.Value == CellValues.Error && cell.CellValue?.InnerText == "#N/A");
            } finally {
                TryDelete(filePath);
            }
        }

        [Fact]
        public void FormulaEvaluator_UsesCustomFunctionsDuringSaveCalculation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.Calculation.RegisterCustomFunction("DOUBLEVALUE", (_, arguments) =>
                        arguments.Count == 1 && arguments[0].Kind == ExcelFormulaValueKind.Number
                            ? ExcelFormulaValue.FromNumber(arguments[0].Number * 2d)
                            : null);
                    document.Calculation.EvaluateFormulasBeforeSave = true;

                    ExcelSheet sheet = document.AddWorksheet("Save Calculation");
                    sheet.CellValue(1, 1, 21d);
                    sheet.CellFormula(1, 2, "DOUBLEVALUE(A1)");
                    document.Save();
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
                Cell formula = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "B1");
                Assert.Equal("DOUBLEVALUE(A1)", formula.CellFormula?.Text);
                Assert.Equal("42", formula.CellValue?.InnerText);
            } finally {
                TryDelete(filePath);
            }
        }

        [Fact]
        public void FormulaEvaluator_RejectsInvalidOrBuiltInCustomFunctionNames() {
            var options = new ExcelCalculationOptions();
            ExcelCustomFormulaFunction function = (_, _) => ExcelFormulaValue.Blank;

            Assert.Throws<ArgumentNullException>(() => options.RegisterCustomFunction("", function));
            Assert.Throws<ArgumentException>(() => options.RegisterCustomFunction("1INVALID", function));
            Assert.Throws<ArgumentException>(() => options.RegisterCustomFunction("BAD-NAME", function));
            Assert.Throws<ArgumentException>(() => options.RegisterCustomFunction("MÉTRIC", function));
            Assert.Throws<ArgumentException>(() => options.RegisterCustomFunction(new string('A', 256), function));
            Assert.Throws<ArgumentException>(() => options.RegisterCustomFunction("SUM", function));
            Assert.Empty(options.CustomFunctionNames);
        }

        [Fact]
        public void FormulaEvaluator_CalculatesExcelFuturePrefixedBuiltInFunctions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    ExcelSheet sheet = document.AddWorksheet("Future Prefixes");
                    sheet.CellValue(1, 1, "East");
                    sheet.CellValue(2, 1, "West");
                    sheet.CellValue(1, 2, 10d);
                    sheet.CellValue(2, 2, 20d);
                    sheet.CellFormula(1, 3, "_xlfn.XLOOKUP(\"West\",A1:A2,B1:B2)");
                    sheet.CellFormula(2, 3, "_xlfn._xlws.XMATCH(\"East\",A1:A2,0)");

                    Assert.Equal(2, sheet.InspectFormulas().SupportedFormulas);
                    Assert.Equal(2, document.Calculate());
                    Assert.Contains(sheet.InspectFormulas().Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "20");
                    Assert.Contains(sheet.InspectFormulas().Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "1");
                    document.Save();
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
                string[] formulas = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<CellFormula>()
                    .Select(formula => formula.Text ?? string.Empty)
                    .ToArray();
                Assert.Contains("_xlfn.XLOOKUP(\"West\",A1:A2,B1:B2)", formulas);
                Assert.Contains("_xlfn._xlws.XMATCH(\"East\",A1:A2,0)", formulas);
            } finally {
                TryDelete(filePath);
            }
        }
    }
}
