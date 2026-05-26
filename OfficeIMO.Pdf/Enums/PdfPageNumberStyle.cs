namespace OfficeIMO.Pdf;

/// <summary>Numbering styles for generated header/footer page tokens.</summary>
public enum PdfPageNumberStyle {
    /// <summary>Arabic numerals: 1, 2, 3.</summary>
    Arabic = 0,
    /// <summary>Lowercase Roman numerals: i, ii, iii.</summary>
    LowerRoman = 1,
    /// <summary>Uppercase Roman numerals: I, II, III.</summary>
    UpperRoman = 2,
    /// <summary>Lowercase alphabetic sequence: a, b, c, aa.</summary>
    LowerLetter = 3,
    /// <summary>Uppercase alphabetic sequence: A, B, C, AA.</summary>
    UpperLetter = 4
}
