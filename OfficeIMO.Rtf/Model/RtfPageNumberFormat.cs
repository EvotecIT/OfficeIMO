namespace OfficeIMO.Rtf;

/// <summary>
/// Page number formats supported by RTF page-numbering controls.
/// </summary>
public enum RtfPageNumberFormat {
    /// <summary>Decimal page numbers, represented by <c>\pgndec</c>.</summary>
    Decimal,

    /// <summary>Uppercase Roman numerals, represented by <c>\pgnucrm</c>.</summary>
    UpperRoman,

    /// <summary>Lowercase Roman numerals, represented by <c>\pgnlcrm</c>.</summary>
    LowerRoman,

    /// <summary>Uppercase letters, represented by <c>\pgnucltr</c>.</summary>
    UpperLetter,

    /// <summary>Lowercase letters, represented by <c>\pgnlcltr</c>.</summary>
    LowerLetter,

    /// <summary>Double-byte decimal numbering, represented by <c>\pgndecd</c>.</summary>
    DoubleByteDecimal
}
