namespace OfficeIMO.Rtf;

/// <summary>
/// Footnote and endnote number formats supported by RTF note-numbering controls.
/// </summary>
public enum RtfNoteNumberFormat {
    /// <summary>Arabic numbers, represented by <c>\ftnnar</c> or <c>\aftnnar</c>.</summary>
    Arabic,

    /// <summary>Lowercase letters, represented by <c>\ftnnalc</c> or <c>\aftnnalc</c>.</summary>
    LowerLetter,

    /// <summary>Uppercase letters, represented by <c>\ftnnauc</c> or <c>\aftnnauc</c>.</summary>
    UpperLetter,

    /// <summary>Lowercase Roman numerals, represented by <c>\ftnnrlc</c> or <c>\aftnnrlc</c>.</summary>
    LowerRoman,

    /// <summary>Uppercase Roman numerals, represented by <c>\ftnnruc</c> or <c>\aftnnruc</c>.</summary>
    UpperRoman
}
