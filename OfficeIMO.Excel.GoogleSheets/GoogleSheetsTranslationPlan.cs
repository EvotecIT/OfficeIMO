using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Pre-export analysis of how a workbook maps to Google Sheets.
    /// </summary>
    public sealed class GoogleSheetsTranslationPlan {
        public GoogleSheetsTranslationPlan(TranslationReport report) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        public TranslationReport Report { get; }
        public int SheetCount { get; internal set; }
        public int TableCount { get; internal set; }
        public int ChartCount { get; internal set; }
        public int PivotTableCount { get; internal set; }
        public int HeaderFooterSheetCount { get; internal set; }
        public int HeaderFooterImageSheetCount { get; internal set; }
        public int NamedRangeCount { get; internal set; }
        public bool HasFormulaTranslationRisk { get; internal set; }
        public bool HasPrintLayoutRisk { get; internal set; }
    }
}
