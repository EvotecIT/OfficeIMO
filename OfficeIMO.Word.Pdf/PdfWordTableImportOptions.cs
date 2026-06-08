using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Options for extracting logical PDF tables into a Word document.
    /// </summary>
    public sealed class PdfWordTableImportOptions {
        /// <summary>
        /// PDF text layout options used when a path, stream, or byte array is loaded directly.
        /// </summary>
        public PdfCore.PdfTextLayoutOptions? LayoutOptions { get; set; }

        /// <summary>
        /// Optional inclusive one-based source page ranges used by direct PDF loading overloads.
        /// </summary>
        public IReadOnlyList<PdfCore.PdfPageRange>? PageRanges { get; set; }

        /// <summary>
        /// Maximum body rows to import per detected table. Values less than or equal to zero import all rows.
        /// </summary>
        public int MaxRows { get; set; }

        /// <summary>
        /// Word table style applied to imported tables.
        /// </summary>
        public WordTableStyle TableStyle { get; set; } = WordTableStyle.TableGrid;

        /// <summary>
        /// When true, a short source paragraph is inserted before each imported table.
        /// </summary>
        public bool IncludeSourceCaptions { get; set; } = true;

        /// <summary>
        /// When true, a page break is inserted between imported PDF tables.
        /// </summary>
        public bool PageBreakBetweenTables { get; set; }

        /// <summary>
        /// When true, tables with inferred column headers repeat the first row at the top of each Word page.
        /// </summary>
        public bool RepeatHeaderRows { get; set; } = true;

        /// <summary>
        /// When true, imported tables are set to 100 percent width and columns are distributed evenly.
        /// </summary>
        public bool FitTablesToPageWidth { get; set; } = true;

        /// <summary>
        /// When true, body cells in inferred numeric PDF columns are right-aligned in the generated Word tables.
        /// </summary>
        public bool AlignNumericColumns { get; set; } = true;

        /// <summary>
        /// Paragraph text written when no tables are detected, keeping the produced document meaningful.
        /// </summary>
        public string EmptyDocumentMessage { get; set; } = "No PDF tables detected.";
    }
}
