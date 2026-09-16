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
        /// <summary>Gets the number of source paragraphs examined.</summary>
        public int ParagraphCount { get; internal set; }
        /// <summary>Gets the number of source tables examined.</summary>
        public int TableCount { get; internal set; }
        /// <summary>Gets the number of nested source tables.</summary>
        public int NestedTableCount { get; internal set; }
        /// <summary>Gets the number of source lists examined.</summary>
        public int ListCount { get; internal set; }
        /// <summary>Gets the number of source images examined.</summary>
        public int ImageCount { get; internal set; }
        /// <summary>Gets the number of source charts examined.</summary>
        public int ChartCount { get; internal set; }
        /// <summary>Gets the number of source comments examined.</summary>
        public int CommentCount { get; internal set; }
        /// <summary>Gets the number of source footnotes examined.</summary>
        public int FootnoteCount { get; internal set; }
        /// <summary>Gets the number of source headers examined.</summary>
        public int HeaderCount { get; internal set; }
        /// <summary>Gets the number of source footers examined.</summary>
        public int FooterCount { get; internal set; }
        /// <summary>Gets the number of source watermarks examined.</summary>
        public int WatermarkCount { get; internal set; }
        /// <summary>Gets the number of SmartArt objects examined.</summary>
        public int SmartArtCount { get; internal set; }
        /// <summary>Gets the number of source shapes examined.</summary>
        public int ShapeCount { get; internal set; }
        /// <summary>Gets the number of text boxes examined.</summary>
        public int TextBoxCount { get; internal set; }
        /// <summary>Gets the number of embedded objects examined.</summary>
        public int EmbeddedObjectCount { get; internal set; }
        /// <summary>Gets the number of structured document tags examined.</summary>
        public int StructuredDocumentTagCount { get; internal set; }
        /// <summary>Gets the number of check-box controls examined.</summary>
        public int CheckBoxCount { get; internal set; }
        /// <summary>Gets the number of date-picker controls examined.</summary>
        public int DatePickerCount { get; internal set; }
        /// <summary>Gets the number of drop-down-list controls examined.</summary>
        public int DropDownListCount { get; internal set; }
        /// <summary>Gets the number of combo-box controls examined.</summary>
        public int ComboBoxCount { get; internal set; }
        /// <summary>Gets the number of equations examined.</summary>
        public int EquationCount { get; internal set; }
        /// <summary>Gets whether source sections need position-aware replay in Docs.</summary>
        public bool RequiresIndexAwareSections { get; internal set; }
        /// <summary>Gets whether the report contains any warning or error.</summary>
        public bool HasFlattenedFeatures => Report.HasWarnings || Report.HasErrors;
    }
}
