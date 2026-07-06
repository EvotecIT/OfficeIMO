using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options controlling first-party OfficeIMO PDF export.
    /// </summary>
    public class PdfSaveOptions {
        /// <summary>
        /// PDF creation options passed to the first-party PDF engine. The options are cloned before export.
        /// </summary>
        public PdfCore.PdfOptions? PdfOptions { get; set; }

        /// <summary>
        /// Optional Word-style font family used as the first-party PDF default font. By default, the family maps to the nearest PDF standard font without embedding installed host fonts.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// When true, first-party Word PDF export may load installed system fonts to embed them into the generated PDF.
        /// Defaults to false so untrusted DOCX content cannot silently copy host font files into exported PDFs.
        /// Explicit font data supplied through <see cref="PdfOptions"/> remains available regardless of this setting.
        /// </summary>
        public bool AllowSystemFontEmbedding { get; set; }

        /// <summary>
        /// Built-in generated-text fallback groups applied by the Word PDF converter when system font embedding is allowed.
        /// Defaults to the recommended preset, but no host font files are embedded unless <see cref="AllowSystemFontEmbedding"/> is true.
        /// </summary>
        public PdfCore.PdfTextFallbackFeatures TextFallbacks { get; set; } = PdfCore.PdfTextFallbackFeatures.Default;

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
        /// Shared conversion report populated alongside <see cref="Warnings"/> for wrapper-friendly diagnostics.
        /// The report is cleared at the start of each export.
        /// </summary>
        public PdfCore.PdfConversionReport ConversionReport { get; } = new PdfCore.PdfConversionReport();

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

        /// <summary>
        /// Applies a high-level export profile by setting the Word PDF options that correspond to that profile.
        /// </summary>
        public PdfSaveOptions UseProfile(PdfCore.PdfExportProfile profile) {
            switch (profile) {
                case PdfCore.PdfExportProfile.Faithful:
                    IncludePageNumbers = true;
                    DefaultTableBorders = false;
                    break;
                case PdfCore.PdfExportProfile.Lightweight:
                    IncludePageNumbers = false;
                    DefaultTableBorders = false;
                    break;
                case PdfCore.PdfExportProfile.PrintReady:
                    IncludePageNumbers = true;
                    DefaultTableBorders = true;
                    break;
                case PdfCore.PdfExportProfile.TextOnly:
                    IncludePageNumbers = false;
                    DefaultTableBorders = false;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unsupported PDF export profile.");
            }

            return this;
        }

        internal void ResetExportState() {
            Warnings.Clear();
            ConversionReport.Clear();
        }
    }
}
