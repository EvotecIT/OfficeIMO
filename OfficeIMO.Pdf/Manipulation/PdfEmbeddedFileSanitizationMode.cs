namespace OfficeIMO.Pdf;

/// <summary>How the sanitizer handles embedded and associated files.</summary>
public enum PdfEmbeddedFileSanitizationMode {
    /// <summary>Remove embedded payload references without retaining decoded bytes in the result.</summary>
    Remove,

    /// <summary>Remove embedded payload references and return decoded attachments in the result for caller-controlled quarantine.</summary>
    Quarantine
}
