namespace OfficeIMO.Rtf;

/// <summary>
/// Timestamp fields in the RTF document information destination supported by the lossless editor.
/// </summary>
public enum RtfDocumentInfoTimestampField {
    /// <summary>Document creation timestamp.</summary>
    Created,

    /// <summary>Document revision timestamp.</summary>
    Revised,

    /// <summary>Document print timestamp.</summary>
    Printed,

    /// <summary>Document backup timestamp.</summary>
    BackedUp
}
