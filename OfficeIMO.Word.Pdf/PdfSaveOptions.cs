using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options controlling first-party OfficeIMO PDF export.
    /// </summary>
    public class PdfSaveOptions {
        private PdfCore.PdfResourcePolicy _resourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic();
        /// <summary>
        /// PDF creation options passed to the first-party PDF engine. The options are cloned before export.
        /// </summary>
        public PdfCore.PdfOptions? PdfOptions { get; set; }

        /// <summary>
        /// Optional Word-style font family used as the first-party PDF default font. When the resource policy allows system fonts, an installed family is embedded; otherwise it maps to the nearest PDF standard font.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>Host-resource policy. Defaults to portable deterministic conversion; callers must explicitly opt in before installed host fonts or external resources are read.</summary>
        public PdfCore.PdfResourcePolicy ResourcePolicy {
            get => _resourcePolicy;
            set => _resourcePolicy = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Built-in generated-text fallback groups applied by the Word PDF converter when system font embedding is allowed.
        /// Defaults to the recommended preset, but no host font files are embedded unless <see cref="ResourcePolicy"/> allows them.
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
        public PdfCore.PdfPageOrientation? Orientation { get; set; }

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

        internal List<PdfExportWarning> Warnings { get; } = new List<PdfExportWarning>();

        internal PdfCore.PdfConversionReport Report { get; } = new PdfCore.PdfConversionReport();

        /// <summary>
        /// Determines whether generated page numbers are rendered when the Word source has no page-number field. Defaults to false.
        /// </summary>
        public bool IncludePageNumbers { get; set; }

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
                    IncludePageNumbers = false;
                    DefaultTableBorders = false;
                    break;
                case PdfCore.PdfExportProfile.Lightweight:
                    IncludePageNumbers = false;
                    DefaultTableBorders = false;
                    break;
                case PdfCore.PdfExportProfile.PrintReady:
                    IncludePageNumbers = false;
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

        internal PdfSaveOptions CloneForConversion() => new() {
            PdfOptions = PdfOptions,
            FontFamily = FontFamily,
            ResourcePolicy = ResourcePolicy.Clone(),
            TextFallbacks = TextFallbacks,
            PageSize = PageSize,
            Margins = Margins,
            Orientation = Orientation,
            DefaultPageSize = DefaultPageSize,
            DefaultOrientation = DefaultOrientation,
            Title = Title,
            Author = Author,
            Subject = Subject,
            Keywords = Keywords,
            IncludePageNumbers = IncludePageNumbers,
            PageNumberFormat = PageNumberFormat,
            DefaultTableBorders = DefaultTableBorders
        };
    }
}
