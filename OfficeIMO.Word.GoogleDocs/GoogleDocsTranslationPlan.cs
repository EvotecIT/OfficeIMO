using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Pre-export analysis of how a Word document maps to Google Docs.
    /// </summary>
    public sealed class GoogleDocsTranslationPlan {
        public GoogleDocsTranslationPlan(TranslationReport report) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        public TranslationReport Report { get; }
        public int SectionCount { get; internal set; }
        public int ParagraphCount { get; internal set; }
        public int TableCount { get; internal set; }
        public int NestedTableCount { get; internal set; }
        public int ListCount { get; internal set; }
        public int ImageCount { get; internal set; }
        public int ChartCount { get; internal set; }
        public int CommentCount { get; internal set; }
        public int FootnoteCount { get; internal set; }
        public int HeaderCount { get; internal set; }
        public int FooterCount { get; internal set; }
        public int WatermarkCount { get; internal set; }
        public int SmartArtCount { get; internal set; }
        public int ShapeCount { get; internal set; }
        public int TextBoxCount { get; internal set; }
        public int EmbeddedObjectCount { get; internal set; }
        public int StructuredDocumentTagCount { get; internal set; }
        public int CheckBoxCount { get; internal set; }
        public int DatePickerCount { get; internal set; }
        public int DropDownListCount { get; internal set; }
        public int ComboBoxCount { get; internal set; }
        public int EquationCount { get; internal set; }
        public bool RequiresIndexAwareSections { get; internal set; }
        public bool HasFlattenedFeatures => Report.HasWarnings || Report.HasErrors;
    }
}
