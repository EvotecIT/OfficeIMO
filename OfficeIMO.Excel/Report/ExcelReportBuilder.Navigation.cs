namespace OfficeIMO.Excel {
    // Navigation helpers for ExcelReportBuilder (sections with anchors, header/footer).
    public sealed partial class ExcelReportBuilder {
        /// <summary>
        /// Inserts a section header and a back-to-top link (explicit A1 link).
        /// </summary>
        public ExcelReportBuilder SectionWithAnchor(string text, string? anchorName = null, bool backToTopLink = true, string backToTopText = "↑ Top") {
            Section(text);
            if (backToTopLink) {
                Sheet.SetInternalLink(_row, 1, Sheet, "A1", backToTopText);
                _row++;
            }
            return this;
        }

        /// <summary>Applies optional autofit operations and returns the composer.</summary>
        public ExcelReportBuilder Finish(bool autoFitColumns = true, bool autoFitRows = false) {
            if (autoFitColumns) Sheet.AutoFitColumns();
            if (autoFitRows) Sheet.AutoFitRows();
            return this;
        }
    }
}
