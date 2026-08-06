using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>Word package type used when creating a document without a destination extension.</summary>
    public enum WordDocumentType {
        /// <summary>Standard DOCX document.</summary>
        Document,

        /// <summary>DOTX template.</summary>
        Template,

        /// <summary>Macro-enabled DOCM document.</summary>
        MacroEnabledDocument,

        /// <summary>Macro-enabled DOTM template.</summary>
        MacroEnabledTemplate
    }

    /// <summary>Controls creation and persistence of a Word document.</summary>
    public sealed class WordCreateOptions : DocumentCreateOptions {
        /// <summary>Controls the Open XML package type when no destination extension is available.</summary>
        public WordDocumentType DocumentType { get; set; } = WordDocumentType.Document;
    }

    /// <summary>Controls access, persistence, and package behavior when loading a Word document.</summary>
    public sealed class WordLoadOptions : DocumentLoadOptions {
        /// <summary>Default maximum complete source size (512 MiB).</summary>
        public const long DefaultMaxInputBytes = 512L * 1024L * 1024L;

        /// <summary>
        /// Maximum number of source bytes accepted before Word package parsing begins.
        /// </summary>
        public long MaxInputBytes { get; set; } = DefaultMaxInputBytes;

        /// <summary>Replaces existing styles with OfficeIMO defaults when the document is editable.</summary>
        public bool OverrideStyles { get; set; }

        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OfficeOpenXmlLoadSettings? OpenSettings { get; set; }
    }

    internal static class WordDocumentTypeExtensions {
        internal static WordprocessingDocumentType ToOpenXml(this WordDocumentType documentType) =>
            documentType switch {
                WordDocumentType.Document => WordprocessingDocumentType.Document,
                WordDocumentType.Template => WordprocessingDocumentType.Template,
                WordDocumentType.MacroEnabledDocument => WordprocessingDocumentType.MacroEnabledDocument,
                WordDocumentType.MacroEnabledTemplate => WordprocessingDocumentType.MacroEnabledTemplate,
                _ => throw new ArgumentOutOfRangeException(nameof(documentType), documentType, "Unsupported Word document type.")
            };
    }
}
