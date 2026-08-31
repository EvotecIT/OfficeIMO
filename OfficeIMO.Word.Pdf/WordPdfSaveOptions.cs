using System.Collections.Generic;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;
using DrawingCore = OfficeIMO.Drawing;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options controlling first-party OfficeIMO PDF export.
    /// </summary>
    public class WordPdfSaveOptions {
        /// <summary>Cancellation observed at document-section and element boundaries during conversion.</summary>
        public CancellationToken CancellationToken { get; set; }

        private PdfCore.PdfResourcePolicy _resourcePolicy = PdfCore.PdfResourcePolicy.CreateDefault();
        private PdfCore.PdfOptions? _pdfOptions;
        private bool _pdfOptionsCreatedByRenderingProfile;
        private long? _renderingProfileFontConfigurationState;
        private long? _renderingProfileFontAssignmentVersion;
        private long? _renderingProfileOwnedFontConfigurationState;
        private long? _renderingProfileOwnedFontAssignmentVersion;
        private long? _renderingProfileDefaultFontSizeAssignmentVersion;
        /// <summary>
        /// PDF creation options passed to the first-party PDF engine. The options are cloned before export.
        /// </summary>
        public PdfCore.PdfOptions? PdfOptions {
            get => _pdfOptions;
            set {
                _pdfOptions = value;
                _pdfOptionsCreatedByRenderingProfile = false;
                _renderingProfileFontConfigurationState = null;
                _renderingProfileFontAssignmentVersion = null;
                _renderingProfileOwnedFontConfigurationState = null;
                _renderingProfileOwnedFontAssignmentVersion = null;
                _renderingProfileDefaultFontSizeAssignmentVersion = null;
            }
        }

        internal bool HasExplicitPdfFontConfiguration =>
            _pdfOptions != null
            && (!_renderingProfileFontConfigurationState.HasValue
                || _pdfOptions.FontConfigurationState
                    != _renderingProfileFontConfigurationState.Value
                || !_renderingProfileFontAssignmentVersion.HasValue
                 || _pdfOptions.FontConfigurationAssignmentVersion
                     != _renderingProfileFontAssignmentVersion.Value);

        internal bool HasExplicitPdfDefaultFontSizeConfiguration =>
            _pdfOptions != null
            && (!_renderingProfileDefaultFontSizeAssignmentVersion.HasValue
                || _pdfOptions.DefaultFontSizeAssignmentVersion
                    != _renderingProfileDefaultFontSizeAssignmentVersion.Value);

        /// <summary>
        /// Optional Word-style font family used as the first-party PDF default font. When the resource policy allows system fonts, an installed family is embedded; otherwise it maps to the nearest PDF standard font.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>Host-resource policy. Defaults to balanced conversion: installed and document-named fonts may be embedded, while local files and remote resources remain disabled.</summary>
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
        public OfficePageOrientation? Orientation { get; set; }

        /// <summary>
        /// Optional default page size applied when creating new documents.
        /// </summary>
        public WordPageSize? DefaultPageSize { get; set; }

        /// <summary>
        /// Optional default page orientation applied when creating new documents.
        /// </summary>
        public OfficePageOrientation? DefaultOrientation { get; set; }

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
        /// Applies shared deterministic typography resources to the first-party PDF engine.
        /// Word pagination and layout settings remain owned by this converter.
        /// </summary>
        public WordPdfSaveOptions UseRenderingProfile(
            DrawingCore.OfficeRenderingProfile profile,
            DrawingCore.OfficeRenderingProfileApplyMode mode = DrawingCore.OfficeRenderingProfileApplyMode.Replace) {
            if (profile == null) {
                throw new ArgumentNullException(nameof(profile));
            }
            if (mode != DrawingCore.OfficeRenderingProfileApplyMode.Replace
                && mode != DrawingCore.OfficeRenderingProfileApplyMode.Overlay) {
                throw new ArgumentOutOfRangeException(nameof(mode));
            }

            bool createdPdfOptions = _pdfOptions == null;
            bool profileOwnsCurrentFontConfiguration =
                _pdfOptions != null
                && _renderingProfileOwnedFontConfigurationState.HasValue
                && _pdfOptions.FontConfigurationState
                    == _renderingProfileOwnedFontConfigurationState.Value
                && _renderingProfileOwnedFontAssignmentVersion.HasValue
                && _pdfOptions.FontConfigurationAssignmentVersion
                    == _renderingProfileOwnedFontAssignmentVersion.Value;
            bool profileOwnsCurrentDefaultFontSize =
                _pdfOptions != null
                && _renderingProfileDefaultFontSizeAssignmentVersion.HasValue
                && _pdfOptions.DefaultFontSizeAssignmentVersion
                    == _renderingProfileDefaultFontSizeAssignmentVersion.Value;
            PdfCore.PdfOptions target = _pdfOptions ?? new PdfCore.PdfOptions();
            target.UseRenderingProfile(profile, mode);
            if (createdPdfOptions) {
                _pdfOptions = target;
                _pdfOptionsCreatedByRenderingProfile = true;
            }
            _renderingProfileFontConfigurationState =
                profile.Fonts.Faces.Count == 0
                && _pdfOptionsCreatedByRenderingProfile
                && (createdPdfOptions
                    || (profileOwnsCurrentFontConfiguration
                        && (mode == DrawingCore.OfficeRenderingProfileApplyMode.Replace
                            || _renderingProfileFontConfigurationState.HasValue)))
                        ? target.FontConfigurationState
                        : null;
            _renderingProfileFontAssignmentVersion =
                _renderingProfileFontConfigurationState.HasValue
                    ? target.FontConfigurationAssignmentVersion
                    : null;
            _renderingProfileOwnedFontConfigurationState =
                _pdfOptionsCreatedByRenderingProfile
                && (createdPdfOptions || profileOwnsCurrentFontConfiguration)
                    ? target.FontConfigurationState
                    : null;
            _renderingProfileOwnedFontAssignmentVersion =
                _renderingProfileOwnedFontConfigurationState.HasValue
                    ? target.FontConfigurationAssignmentVersion
                    : null;
            _renderingProfileDefaultFontSizeAssignmentVersion =
                _pdfOptionsCreatedByRenderingProfile
                && (createdPdfOptions || profileOwnsCurrentDefaultFontSize)
                    ? target.DefaultFontSizeAssignmentVersion
                    : null;
            return this;
        }

        /// <summary>
        /// Applies a high-level export profile by setting the Word PDF options that correspond to that profile.
        /// </summary>
        public WordPdfSaveOptions UseProfile(PdfCore.PdfExportProfile profile) {
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

        internal WordPdfSaveOptions CloneForConversion() {
            var clone = new WordPdfSaveOptions {
                CancellationToken = CancellationToken,
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
            clone._pdfOptionsCreatedByRenderingProfile =
                _pdfOptionsCreatedByRenderingProfile;
            clone._renderingProfileFontConfigurationState =
                _renderingProfileFontConfigurationState;
            clone._renderingProfileFontAssignmentVersion =
                _renderingProfileFontAssignmentVersion;
            clone._renderingProfileOwnedFontConfigurationState =
                _renderingProfileOwnedFontConfigurationState;
            clone._renderingProfileOwnedFontAssignmentVersion =
                _renderingProfileOwnedFontAssignmentVersion;
            clone._renderingProfileDefaultFontSizeAssignmentVersion =
                _renderingProfileDefaultFontSizeAssignmentVersion;
            return clone;
        }
    }
}
