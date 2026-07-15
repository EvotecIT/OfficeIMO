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
                AddUnsupported(report, "Charts", "SHEETS.CHART", plan.ChartCount, options.UnsupportedFeatures.Charts, "charts");
            }

            if (plan.PivotTableCount > 0) {
                AddUnsupported(report, "PivotTables", "SHEETS.PIVOT_TABLE", plan.PivotTableCount, options.UnsupportedFeatures.PivotTables, "pivot tables");
            }

            if (plan.HeaderFooterSheetCount > 0) {
                plan.HasPrintLayoutRisk = true;
                AddUnsupported(report, "PrintLayout", "SHEETS.PRINT_LAYOUT", plan.HeaderFooterSheetCount, options.UnsupportedFeatures.PrintLayout, "worksheets with print header/footer metadata");
            }

            if (plan.HeaderFooterImageSheetCount > 0) {
                report.Add(TranslationSeverity.Warning, "HeaderFooterImages", "Header/footer images are present and should be considered unsupported until a Google Sheets path is verified.");
            }

            plan.HasFormulaTranslationRisk = true;
            report.Add(TranslationSeverity.Info, "Formulas", "Formula compatibility needs an Excel-to-Google function mapping and a diagnostic path for unsupported formulas.");

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
