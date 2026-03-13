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

            if (plan.ChartCount > 0 && !options.IncludeCharts) {
                report.Add(TranslationSeverity.Warning, "Charts", "Charts are present but chart export is disabled in the current plan.");
            } else if (plan.ChartCount > 0) {
                report.Add(TranslationSeverity.Info, "Charts", "Charts are present and will require chart-spec translation to Google Sheets.");
            }

            if (plan.PivotTableCount > 0 && !options.IncludePivotTables) {
                report.Add(TranslationSeverity.Warning, "PivotTables", "Pivot tables are present but pivot export is disabled in the current plan.");
            } else if (plan.PivotTableCount > 0) {
                report.Add(TranslationSeverity.Info, "PivotTables", "Pivot tables are present and will require dedicated Google Sheets pivot mapping.");
            }

            if (plan.HeaderFooterSheetCount > 0) {
                plan.HasPrintLayoutRisk = true;
                var message = options.TreatPrintLayoutAsDiagnosticOnly
                    ? "Header/footer metadata exists and should be treated as diagnostic-only until a concrete Google Sheets print-layout mapping is confirmed."
                    : "Header/footer metadata exists and needs a dedicated Google Sheets print-layout mapping strategy.";
                report.Add(TranslationSeverity.Warning, "PrintLayout", message);
            }

            if (plan.HeaderFooterImageSheetCount > 0) {
                report.Add(TranslationSeverity.Warning, "HeaderFooterImages", "Header/footer images are present and should be considered unsupported until a Google Sheets path is verified.");
            }

            plan.HasFormulaTranslationRisk = true;
            report.Add(TranslationSeverity.Info, "Formulas", "Formula compatibility needs an Excel-to-Google function mapping and a diagnostic path for unsupported formulas.");

            report.Add(TranslationSeverity.Info, "NamedRanges", "Named range counting is left conservative for now because the current public workbook API does not expose a named-range enumerator.");

            return plan;
        }
    }
}
