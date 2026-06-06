namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Describes workbook content that could not be mapped faithfully during Excel-to-PDF export.
    /// </summary>
    public sealed class ExcelPdfExportWarning {
        /// <summary>Creates a warning for content skipped or simplified during export.</summary>
        public ExcelPdfExportWarning(string sheetName, string feature, string message) {
            SheetName = sheetName ?? string.Empty;
            Feature = feature ?? string.Empty;
            Message = message ?? string.Empty;
        }

        /// <summary>Worksheet name associated with the warning.</summary>
        public string SheetName { get; }

        /// <summary>Feature area associated with the warning.</summary>
        public string Feature { get; }

        /// <summary>Human-readable warning message.</summary>
        public string Message { get; }

        /// <inheritdoc />
        public override string ToString() {
            if (string.IsNullOrWhiteSpace(SheetName)) {
                return string.IsNullOrWhiteSpace(Feature) ? Message : Feature + ": " + Message;
            }

            return string.IsNullOrWhiteSpace(Feature)
                ? SheetName + ": " + Message
                : SheetName + " [" + Feature + "]: " + Message;
        }

        /// <summary>
        /// Converts this Excel-specific warning to the shared PDF conversion warning contract.
        /// </summary>
        public OfficeIMO.Pdf.PdfConversionWarning ToConversionWarning() {
            var details = new Dictionary<string, string>();
            if (!string.IsNullOrWhiteSpace(SheetName)) {
                details["SheetName"] = SheetName;
            }

            if (!string.IsNullOrWhiteSpace(Feature)) {
                details["Feature"] = Feature;
            }

            return new OfficeIMO.Pdf.PdfConversionWarning(
                "OfficeIMO.Excel.Pdf",
                string.IsNullOrWhiteSpace(Feature) ? "ExcelPdfWarning" : Feature,
                SheetName,
                Message,
                details: details);
        }
    }
}
