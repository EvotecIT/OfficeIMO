using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options for importing parser-supported PDF content into an editable Word document.
    /// </summary>
    /// <remarks>
    /// PDF import is semantic reconstruction over the first-party logical PDF reader. It preserves
    /// supported metadata, page breaks, headings, paragraphs, list items, logical tables, and source
    /// placeholders, but it is not a pixel-perfect fixed-layout PDF to DOCX renderer.
    /// </remarks>
    public sealed class PdfWordReadOptions {
        /// <summary>Logical PDF layout options used while grouping positioned PDF text into semantic objects.</summary>
        public PdfCore.PdfTextLayoutOptions? LayoutOptions { get; set; }

        /// <summary>Optional inclusive one-based source page ranges used by direct PDF loading overloads.</summary>
        public IReadOnlyList<PdfCore.PdfPageRange>? PageRanges { get; set; }

        /// <summary>Whether PDF Info dictionary metadata should be copied into Word built-in properties.</summary>
        public bool IncludeMetadata { get; set; } = true;

        /// <summary>Whether source PDF page transitions should be represented by Word page breaks.</summary>
        public bool PreservePageBreaks { get; set; } = true;

        /// <summary>Whether empty PDF pages should produce an empty Word paragraph when page breaks are preserved.</summary>
        public bool IncludeEmptyPages { get; set; }

        /// <summary>Whether logical heading lines should be imported as Word heading paragraphs.</summary>
        public bool ImportHeadings { get; set; } = true;

        /// <summary>Whether grouped logical paragraphs should be imported as Word paragraphs.</summary>
        public bool ImportParagraphs { get; set; } = true;

        /// <summary>Whether logical list items should be imported as editable Word list items.</summary>
        public bool ImportLists { get; set; } = true;

        /// <summary>Whether logical PDF tables should be imported as editable Word tables.</summary>
        public bool ImportTables { get; set; } = true;

        /// <summary>Whether safe URI link annotations should be imported as editable Word hyperlinks.</summary>
        public bool ImportUriLinks { get; set; } = true;

        /// <summary>Whether supported PDF internal destination links should be imported as Word bookmark hyperlinks.</summary>
        public bool ImportInternalLinks { get; set; } = true;

        /// <summary>Prefix used for generated Word bookmarks that represent imported PDF pages and named destinations.</summary>
        public string BookmarkPrefix { get; set; } = "OfficeIMO_Pdf";

        /// <summary>Absolute URI schemes allowed when creating active Word hyperlinks from PDF URI annotations.</summary>
        public ISet<string> AllowedHyperlinkUriSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            Uri.UriSchemeHttp,
            Uri.UriSchemeHttps,
            Uri.UriSchemeMailto
        };

        /// <summary>Whether complete PDF image-file payloads should be embedded as native Word images.</summary>
        public bool ImportImages { get; set; } = true;

        /// <summary>Whether embedded images should use detected PDF placement size when available.</summary>
        public bool PreserveImagePlacementSize { get; set; } = true;

        /// <summary>Whether image resources should be represented by editable placeholder paragraphs when native image embedding is unavailable or disabled.</summary>
        public bool IncludeImagePlaceholders { get; set; } = true;

        /// <summary>Whether AcroForm widgets should be represented by editable placeholder paragraphs.</summary>
        public bool IncludeFormFieldPlaceholders { get; set; } = true;

        /// <summary>Maximum body rows to import per detected table. Values less than or equal to zero import all rows.</summary>
        public int MaxTableRows { get; set; }

        /// <summary>Word table style applied to imported PDF tables.</summary>
        public WordTableStyle TableStyle { get; set; } = WordTableStyle.TableGrid;

        /// <summary>When true, tables with inferred column headers repeat the first row at the top of each Word page.</summary>
        public bool RepeatHeaderRows { get; set; } = true;

        /// <summary>When true, imported tables are set to 100 percent width and columns are distributed evenly.</summary>
        public bool FitTablesToPageWidth { get; set; } = true;

        /// <summary>When true, body cells in inferred numeric PDF columns are right-aligned in the generated Word tables.</summary>
        public bool AlignNumericColumns { get; set; } = true;

        /// <summary>Paragraph text written when no supported PDF content is detected, keeping the produced document meaningful.</summary>
        public string EmptyDocumentMessage { get; set; } = "No supported PDF content detected.";

        /// <summary>Shared conversion report populated with accepted-degradation diagnostics for this import run.</summary>
        public PdfCore.PdfConversionReport ConversionReport { get; } = new PdfCore.PdfConversionReport();

        /// <summary>Creates a reusable copy of this option set.</summary>
        public PdfWordReadOptions Clone() => new PdfWordReadOptions {
            LayoutOptions = CloneLayoutOptions(LayoutOptions),
            PageRanges = PageRanges == null ? null : PageRanges.ToArray(),
            IncludeMetadata = IncludeMetadata,
            PreservePageBreaks = PreservePageBreaks,
            IncludeEmptyPages = IncludeEmptyPages,
            ImportHeadings = ImportHeadings,
            ImportParagraphs = ImportParagraphs,
            ImportLists = ImportLists,
            ImportTables = ImportTables,
            ImportUriLinks = ImportUriLinks,
            ImportInternalLinks = ImportInternalLinks,
            BookmarkPrefix = BookmarkPrefix,
            ImportImages = ImportImages,
            PreserveImagePlacementSize = PreserveImagePlacementSize,
            IncludeImagePlaceholders = IncludeImagePlaceholders,
            IncludeFormFieldPlaceholders = IncludeFormFieldPlaceholders,
            MaxTableRows = MaxTableRows,
            TableStyle = TableStyle,
            RepeatHeaderRows = RepeatHeaderRows,
            FitTablesToPageWidth = FitTablesToPageWidth,
            AlignNumericColumns = AlignNumericColumns,
            EmptyDocumentMessage = EmptyDocumentMessage
        }.CopyAllowedHyperlinkUriSchemesFrom(AllowedHyperlinkUriSchemes);

        internal void ResetImportState() {
            ConversionReport.Clear();
        }

        private static PdfCore.PdfTextLayoutOptions? CloneLayoutOptions(PdfCore.PdfTextLayoutOptions? options) {
            if (options is null) {
                return null;
            }

            return new PdfCore.PdfTextLayoutOptions {
                MarginLeft = options.MarginLeft,
                MarginRight = options.MarginRight,
                BinWidth = options.BinWidth,
                MinGutterWidth = options.MinGutterWidth,
                LineMergeToleranceEm = options.LineMergeToleranceEm,
                LineMergeMaxPoints = options.LineMergeMaxPoints,
                ForceSingleColumn = options.ForceSingleColumn,
                JoinHyphenationAcrossLines = options.JoinHyphenationAcrossLines,
                IgnoreHeaderHeight = options.IgnoreHeaderHeight,
                IgnoreFooterHeight = options.IgnoreFooterHeight,
                GapSpaceThresholdEm = options.GapSpaceThresholdEm,
                GapGlyphFactor = options.GapGlyphFactor
            };
        }

        private PdfWordReadOptions CopyAllowedHyperlinkUriSchemesFrom(IEnumerable<string> schemes) {
            AllowedHyperlinkUriSchemes.Clear();
            foreach (string scheme in schemes) {
                if (!string.IsNullOrWhiteSpace(scheme)) {
                    AllowedHyperlinkUriSchemes.Add(scheme);
                }
            }

            return this;
        }
    }
}
