using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static class GoogleSheetsPlanBuilder {
        internal static GoogleSheetsTranslationPlan Build(ExcelDocument document, GoogleSheetsSaveOptions options) {
            var report = new TranslationReport();
            var plan = new GoogleSheetsTranslationPlan(report) {
                SheetCount = document.Sheets.Count,
                TableCount = document.GetTables().Count,
                PivotTableCount = document.GetPivotTables().Count,
                NamedRangeCount = 0,
            };
            var unsupportedFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var inspection = document.CreateInspectionSnapshot(new ExcelReadOptions {
                UseCachedFormulaResult = true,
                TreatDatesUsingNumberFormat = true,
            });
            foreach (ExcelCellSnapshot cell in inspection.Worksheets.SelectMany(sheet => sheet.Cells)) {
                if (string.IsNullOrWhiteSpace(cell.Formula)) continue;
                plan.FormulaCount++;
                GoogleSheetsFormulaTranslation translation = GoogleSheetsFormulaCatalog.Translate(cell.Formula!, options.Formulas);
                if (!translation.IsSupported) {
                    plan.UnsupportedFormulaCount++;
                    foreach (string function in translation.UnsupportedFunctions) unsupportedFunctions.Add(function);
                }
            }

            foreach (var sheet in document.Sheets) {
                plan.ChartCount += sheet.Charts.Count();

                var headerFooter = sheet.GetHeaderFooter();
                if (!string.IsNullOrEmpty(headerFooter.HeaderLeft)
                    || !string.IsNullOrEmpty(headerFooter.HeaderCenter)
                    || !string.IsNullOrEmpty(headerFooter.HeaderRight)
                    || !string.IsNullOrEmpty(headerFooter.FooterLeft)
                    || !string.IsNullOrEmpty(headerFooter.FooterCenter)
                    || !string.IsNullOrEmpty(headerFooter.FooterRight)
                    || headerFooter.DifferentFirstPage
                    || headerFooter.DifferentOddEven) {
                    plan.HeaderFooterSheetCount++;
                }

                if (headerFooter.HeaderHasPicturePlaceholder || headerFooter.FooterHasPicturePlaceholder) {
                    plan.HeaderFooterImageSheetCount++;
                }
            }

            if (plan.ChartCount > 0) {
                report.Add(TranslationSeverity.Info, "Charts", $"{plan.ChartCount} chart(s) will be evaluated against the code-owned chart support matrix.",
                    code: "SHEETS.CHART.PREFLIGHT", action: TranslationAction.Preserve, count: plan.ChartCount);
            }

            if (plan.PivotTableCount > 0) {
                report.Add(TranslationSeverity.Info, "PivotTables", $"{plan.PivotTableCount} pivot table(s) will be evaluated against the code-owned pivot support matrix.",
                    code: "SHEETS.PIVOT_TABLE.PREFLIGHT", action: TranslationAction.Preserve, count: plan.PivotTableCount);
            }

            if (plan.HeaderFooterSheetCount > 0) {
                plan.HasPrintLayoutRisk = true;
                AddUnsupported(report, "PrintLayout", "SHEETS.PRINT_LAYOUT", plan.HeaderFooterSheetCount, options.UnsupportedFeatures.PrintLayout, "worksheets with print header/footer metadata");
            }

            if (plan.HeaderFooterImageSheetCount > 0) {
                report.Add(TranslationSeverity.Warning, "HeaderFooterImages", "Header/footer images are present and should be considered unsupported until a Google Sheets path is verified.");
            }

            plan.HasFormulaTranslationRisk = plan.UnsupportedFormulaCount > 0;
            if (plan.FormulaCount > 0 && plan.UnsupportedFormulaCount == 0) {
                report.Add(
                    TranslationSeverity.Info,
                    "Formulas",
                    $"All {plan.FormulaCount} formulas use functions recognized by the OfficeIMO Google Sheets compatibility catalog.",
                    code: "SHEETS.FORMULA.SUPPORTED",
                    action: TranslationAction.Preserve,
                    count: plan.FormulaCount);
            } else if (plan.UnsupportedFormulaCount > 0) {
                TranslationSeverity severity = options.Formulas.UnsupportedFormulaMode == GoogleSheetsUnsupportedFormulaMode.Error
                    ? TranslationSeverity.Error
                    : TranslationSeverity.Warning;
                TranslationAction action = options.Formulas.UnsupportedFormulaMode == GoogleSheetsUnsupportedFormulaMode.UseCachedValue
                    ? TranslationAction.Flatten
                    : options.Formulas.UnsupportedFormulaMode == GoogleSheetsUnsupportedFormulaMode.Error
                        ? TranslationAction.Fail
                        : TranslationAction.Preserve;
                report.Add(
                    severity,
                    "Formulas",
                    $"{plan.UnsupportedFormulaCount} formulas use functions not in the compatibility catalog: {string.Join(", ", unsupportedFunctions.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))}.",
                    code: "SHEETS.FORMULA.UNSUPPORTED",
                    action: action,
                    count: plan.UnsupportedFormulaCount);
            }

            report.Add(TranslationSeverity.Info, "NamedRanges", "Named range counting is left conservative for now because the current public workbook API does not expose a named-range enumerator.");

            return plan;
        }

        private static void AddUnsupported(
            TranslationReport report,
            string feature,
            string codePrefix,
            int count,
            UnsupportedFeatureMode mode,
            string description) {
            switch (mode) {
                case UnsupportedFeatureMode.Error:
                    report.Add(
                        TranslationSeverity.Error,
                        feature,
                        $"The workbook contains {count} {description}, and the selected policy requires native preservation.",
                        code: codePrefix + ".UNSUPPORTED",
                        action: TranslationAction.Fail,
                        count: count);
                    break;
                case UnsupportedFeatureMode.WarnAndSkip:
                    report.Add(
                        TranslationSeverity.Warning,
                        feature,
                        $"The workbook contains {count} {description}; the current Google Sheets exporter will skip them.",
                        code: codePrefix + ".SKIPPED",
                        action: TranslationAction.Skip,
                        count: count);
                    break;
                case UnsupportedFeatureMode.Flatten:
                case UnsupportedFeatureMode.Rasterize:
                    report.Add(
                        TranslationSeverity.Error,
                        feature,
                        $"The selected {mode} policy for {description} cannot execute because no compatible adapter is configured.",
                        code: codePrefix + ".FALLBACK_UNAVAILABLE",
                        action: TranslationAction.Fail,
                        count: count);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(mode));
            }
        }
    }
}
