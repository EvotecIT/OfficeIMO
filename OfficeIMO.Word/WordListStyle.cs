namespace OfficeIMO.Word;

/// <summary>
/// List of styles for Word lists
/// Most of the styles are based on the built-in Word list styles, except for the Custom style
/// </summary>
public enum WordListStyle {
    /// <summary>
    /// Regular bulleted list.
    /// </summary>
    Bulleted = 0,
    /// <summary>
    /// Numbered style used for article sections.
    /// </summary>
    ArticleSections = 1,
    /// <summary>
    /// Three-level heading numbering (1.1.1).
    /// </summary>
    Headings111 = 2,
    /// <summary>
    /// Multi-level heading style starting with Roman numerals.
    /// </summary>
    HeadingIA1 = 3,
    /// <summary>
    /// Chapter numbering style.
    /// </summary>
    Chapters = 4,
    /// <summary>
    /// Bulleted list using Wingdings characters.
    /// </summary>
    BulletedChars = 5,
    /// <summary>
    /// Heading style using A.i numbering.
    /// </summary>
    Heading1ai = 6,
    /// <summary>
    /// Variation of Headings111 with shifted levels.
    /// </summary>
    Headings111Shifted = 7,
    /// <summary>
    /// Lowercase letters followed by a bracket.
    /// </summary>
    LowerLetterWithBracket = 8,
    /// <summary>
    /// Lowercase letters followed by a dot.
    /// </summary>
    LowerLetterWithDot = 9,
    /// <summary>
    /// Uppercase letters followed by a dot.
    /// </summary>
    UpperLetterWithDot = 10,
    /// <summary>
    /// Uppercase letters followed by a bracket.
    /// </summary>
    UpperLetterWithBracket = 11,
    /// <summary>
    /// Custom numbering defined by the user.
    /// </summary>
    Custom = 12,
    /// <summary>
    /// Standard numbered list.
    /// </summary>
    Numbered = 13
}
