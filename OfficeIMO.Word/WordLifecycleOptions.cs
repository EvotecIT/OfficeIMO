using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>Controls creation and persistence of a Word document.</summary>
    public sealed class WordCreateOptions : DocumentCreateOptions {
        /// <summary>Controls the Open XML package type when no destination extension is available.</summary>
        public WordprocessingDocumentType DocumentType { get; set; } = WordprocessingDocumentType.Document;
    }

    /// <summary>Controls access, persistence, and package behavior when loading a Word document.</summary>
    public sealed class WordLoadOptions : DocumentLoadOptions {
        /// <summary>Replaces existing styles with OfficeIMO defaults when the document is editable.</summary>
        public bool OverrideStyles { get; set; }

        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OpenSettings? OpenSettings { get; set; }
    }
}
