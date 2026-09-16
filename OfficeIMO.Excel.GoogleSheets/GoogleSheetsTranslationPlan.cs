using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Pre-export analysis of how a workbook maps to Google Sheets.
    /// </summary>
    public sealed class GoogleSheetsTranslationPlan {
        /// <summary>Creates a plan associated with a translation report.</summary>
        public GoogleSheetsTranslationPlan(TranslationReport report) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets fidelity and preflight notices from planning.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets the number of source worksheets examined.</summary>
        public int SheetCount { get; internal set; }
        /// <summary>Gets the number of source tables examined.</summary>
        public int TableCount { get; internal set; }
        /// <summary>Gets the number of source charts examined.</summary>
        public int ChartCount { get; internal set; }
        /// <summary>Gets the number of source pivot tables examined.</summary>
        public int PivotTableCount { get; internal set; }
        /// <summary>Gets the number of sheets with header or footer content.</summary>
        public int HeaderFooterSheetCount { get; internal set; }
        /// <summary>Gets the number of sheets with header or footer images.</summary>
        public int HeaderFooterImageSheetCount { get; internal set; }
        /// <summary>Gets zero because planning does not yet enumerate source defined names; zero does not imply that the workbook has no defined names.</summary>
        public int NamedRangeCount { get; internal set; }
        /// <summary>Gets whether any formula uses a function classified as unsupported.</summary>
        public bool HasFormulaTranslationRisk { get; internal set; }
        /// <summary>Gets the number of source formula cells examined.</summary>
        public int FormulaCount { get; internal set; }
        /// <summary>Gets the number of formula cells using unsupported functions.</summary>
        public int UnsupportedFormulaCount { get; internal set; }
        /// <summary>Gets whether source print-layout features have no equivalent.</summary>
        public bool HasPrintLayoutRisk { get; internal set; }
    }
}
