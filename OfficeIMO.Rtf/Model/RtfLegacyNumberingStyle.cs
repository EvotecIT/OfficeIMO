namespace OfficeIMO.Rtf;

/// <summary>
/// Number text style for Word 6/95 legacy RTF paragraph numbering.
/// </summary>
public enum RtfLegacyNumberingStyle {
    /// <summary>No explicit legacy numbering style.</summary>
    None,

    /// <summary>Cardinal text numbering represented by <c>\pncard</c>.</summary>
    Cardinal,

    /// <summary>Decimal numbering represented by <c>\pndec</c>.</summary>
    Decimal,

    /// <summary>Uppercase alphabetic numbering represented by <c>\pnucltr</c>.</summary>
    UpperLetter,

    /// <summary>Uppercase Roman numbering represented by <c>\pnucrm</c>.</summary>
    UpperRoman,

    /// <summary>Lowercase alphabetic numbering represented by <c>\pnlcltr</c>.</summary>
    LowerLetter,

    /// <summary>Lowercase Roman numbering represented by <c>\pnlcrm</c>.</summary>
    LowerRoman,

    /// <summary>Ordinal numbering represented by <c>\pnord</c>.</summary>
    Ordinal,

    /// <summary>Ordinal text numbering represented by <c>\pnordt</c>.</summary>
    OrdinalText
}
