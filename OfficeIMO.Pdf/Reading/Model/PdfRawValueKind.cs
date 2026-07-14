namespace OfficeIMO.Pdf;

/// <summary>Kind of an immutable raw PDF syntax projection.</summary>
public enum PdfRawValueKind {
    /// <summary>PDF null value.</summary>
    Null,
    /// <summary>Numeric scalar.</summary>
    Number,
    /// <summary>Boolean scalar.</summary>
    Boolean,
    /// <summary>PDF name value.</summary>
    Name,
    /// <summary>PDF literal or hexadecimal string value.</summary>
    TextString,
    /// <summary>PDF array.</summary>
    Array,
    /// <summary>PDF dictionary.</summary>
    Dictionary,
    /// <summary>Indirect object reference.</summary>
    Reference,
    /// <summary>Stream dictionary and length metadata.</summary>
    Stream,
    /// <summary>Value omitted because a projection limit was reached.</summary>
    Truncated
}
