using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Pre-export analysis of how a Word document maps to Google Docs.
    /// </summary>
    public sealed class GoogleDocsTranslationPlan {
        /// <summary>Creates a plan associated with its translation report.</summary>
        public GoogleDocsTranslationPlan(TranslationReport report) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets fidelity and preflight notices from planning.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets the number of source sections examined.</summary>
        public int SectionCount { get; internal set; }
        /// <summary>Gets the number of top-level body paragraphs in the source; paragraphs inside tables, headers, and footers are not included.</summary>
        public int ParagraphCount { get; internal set; }
        /// <summary>Gets the number of top-level source tables; nested tables are counted separately.</summary>
        public int TableCount { get; internal set; }
        /// <summary>Gets the additional nested source tables beyond the top-level table count.</summary>
        public int NestedTableCount { get; internal set; }
        /// <summary>Gets the number of source lists in the document's list collection.</summary>
        public int ListCount { get; internal set; }
        /// <summary>Gets the number of source images in the document image collection.</summary>
        public int ImageCount { get; internal set; }
        /// <summary>Gets the number of source charts in the document chart collection.</summary>
        public int ChartCount { get; internal set; }
        /// <summary>Gets the number of source comments in the document comment collection.</summary>
        public int CommentCount { get; internal set; }
        /// <summary>Gets the number of source footnotes in the document footnote collection.</summary>
        public int FootnoteCount { get; internal set; }
        /// <summary>Gets the number of configured default, even, and first-page header variants across sections.</summary>
        public int HeaderCount { get; internal set; }
        /// <summary>Gets the number of configured default, even, and first-page footer variants across sections.</summary>
        public int FooterCount { get; internal set; }
        /// <summary>Gets the number of source watermarks in the document watermark collection.</summary>
        public int WatermarkCount { get; internal set; }
        /// <summary>Gets the number of source SmartArt objects in the document collection.</summary>
        public int SmartArtCount { get; internal set; }
        /// <summary>Gets the number of source shapes in the document shape collection.</summary>
        public int ShapeCount { get; internal set; }
        /// <summary>Gets the number of source text boxes in the document text-box collection.</summary>
        public int TextBoxCount { get; internal set; }
        /// <summary>Gets the number of source embedded objects in the document collection.</summary>
        public int EmbeddedObjectCount { get; internal set; }
        /// <summary>Gets the number of source structured document tags in the document collection.</summary>
        public int StructuredDocumentTagCount { get; internal set; }
        /// <summary>Gets the number of source check-box controls in the document collection.</summary>
        public int CheckBoxCount { get; internal set; }
        /// <summary>Gets the number of source date-picker controls in the document collection.</summary>
        public int DatePickerCount { get; internal set; }
        /// <summary>Gets the number of source drop-down-list controls in the document collection.</summary>
        public int DropDownListCount { get; internal set; }
        /// <summary>Gets the number of source combo-box controls in the document collection.</summary>
        public int ComboBoxCount { get; internal set; }
        /// <summary>Gets the number of equations exposed by the source paragraph-equation collection.</summary>
        public int EquationCount { get; internal set; }
        /// <summary>Gets whether source sections need position-aware replay in Docs.</summary>
        public bool RequiresIndexAwareSections { get; internal set; }
        /// <summary>Gets whether the report contains any warning or error.</summary>
        public bool HasFlattenedFeatures => Report.HasWarnings || Report.HasErrors;
    }
}
