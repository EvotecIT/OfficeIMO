using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options controlling PDF export.
    /// </summary>
    public class PdfSaveOptions : ConversionOptions {
        /// <summary>
        /// Optional page size for the generated PDF.
        /// </summary>
        public PageSize? PageSize { get; set; }

        /// <summary>
        /// Optional page orientation for the generated PDF.
        /// </summary>
        public PdfPageOrientation? Orientation { get; set; }

        /// <summary>
        /// Optional page margins for the generated PDF.
        /// </summary>
        public float? Margin { get; set; }

        /// <summary>
        /// Measurement unit for the margin value.
        /// </summary>
        public Unit MarginUnit { get; set; } = Unit.Centimetre;
    }
}
