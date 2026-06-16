namespace OfficeIMO.Rtf;

/// <summary>
/// RTF document information fields supported by the lossless editor.
/// </summary>
public enum RtfDocumentInfoField {
    /// <summary>Document title.</summary>
    Title,

    /// <summary>Document subject.</summary>
    Subject,

    /// <summary>Document author.</summary>
    Author,

    /// <summary>Document manager.</summary>
    Manager,

    /// <summary>Document company.</summary>
    Company,

    /// <summary>Document operator.</summary>
    Operator,

    /// <summary>Document category.</summary>
    Category,

    /// <summary>Document keywords.</summary>
    Keywords,

    /// <summary>Document comments.</summary>
    Comments,

    /// <summary>Hyperlink base address.</summary>
    HyperlinkBase
}
