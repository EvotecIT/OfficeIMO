using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options controlling first-party OfficeIMO PDF export.
    /// </summary>
    public class PdfSaveOptions {
        /// <summary>
        /// Optional Word-style font family used as the first-party PDF default font when it maps to Helvetica, Times, or Courier standard PDF families.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// Optional page size in PDF points. The supplied geometry is preserved unless <see cref="Orientation"/> is also set.
        /// </summary>
        public PdfCore.PageSize? PageSize { get; set; }

        /// <summary>
        /// Optional page margins in PDF points.
        /// </summary>
        public PdfCore.PageMargins? Margins { get; set; }

        /// <summary>
        /// Optional page orientation for the generated PDF.
        /// </summary>
        public PdfPageOrientation? Orientation { get; set; }

        /// <summary>
        /// Optional default page size applied when creating new documents.
        /// </summary>
        public WordPageSize? DefaultPageSize { get; set; }

        /// <summary>
        /// Optional default page orientation applied when creating new documents.
        /// </summary>
        public DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues? DefaultOrientation { get; set; }

        /// <summary>
        /// Optional PDF title that overrides the Word document title.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Optional PDF author that overrides the Word document author.
        /// </summary>
        public string? Author { get; set; }

        /// <summary>
        /// Optional PDF subject that overrides the Word document subject.
        /// </summary>
        public string? Subject { get; set; }

        /// <summary>
        /// Optional PDF keywords that override the Word document keywords.
        /// </summary>
        public string? Keywords { get; set; }

        /// <summary>
        /// Warnings populated when content cannot be mapped faithfully.
        /// The collection is cleared at the start of each export.
        /// </summary>
        public List<PdfExportWarning> Warnings { get; } = new List<PdfExportWarning>();

        /// <summary>
        /// Determines whether page numbers are rendered in the PDF footer. Defaults to true.
        /// </summary>
        public bool IncludePageNumbers { get; set; } = true;

        /// <summary>
        /// Optional format for page numbers. Use "{current}" for the current page and "{total}" for total pages.
        /// </summary>
        public string? PageNumberFormat { get; set; }

        /// <summary>
        /// When true, draws subtle borders for table cells that do not define borders in the Word document.
        /// Defaults to false to preserve strict fidelity.
        /// </summary>
        public bool DefaultTableBorders { get; set; } = false;
    }
}
