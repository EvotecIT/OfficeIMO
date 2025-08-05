using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options controlling PDF export.
    /// </summary>
    public class PdfSaveOptions {
        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string FontFamily { get; set; }
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
        
        /// <summary>
        /// Optional default page size applied when creating new documents.
        /// </summary>
        public WordPageSize? DefaultPageSize { get; set; }
        
        /// <summary>
        /// Optional default page orientation applied when creating new documents.
        /// </summary>
        public DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues? DefaultOrientation { get; set; }
    }
}
