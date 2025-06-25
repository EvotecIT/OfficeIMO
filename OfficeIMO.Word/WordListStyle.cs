using System;

namespace OfficeIMO.Word;

/// <summary>
/// List of styles for Word lists
/// Most of the styles are based on the built-in Word list styles, except for the Custom style
/// </summary>
public enum WordListStyle {
    /// <summary>
    /// Regular bulleted list.
    /// </summary>
    Bulleted,
    /// <summary>
    /// Numbered style used for article sections.
    /// </summary>
    ArticleSections,
    /// <summary>
    /// Three-level heading numbering (1.1.1).
    /// </summary>
    Headings111,
    /// <summary>
    /// Multi-level heading style starting with Roman numerals.
    /// </summary>
    HeadingIA1,
    /// <summary>
    /// Chapter numbering style.
    /// </summary>
    Chapters,
    /// <summary>
    /// Bulleted list using Wingdings characters.
    /// </summary>
    BulletedChars,
    /// <summary>
    /// Heading style using A.i numbering.
    /// </summary>
    Heading1ai,
    /// <summary>
    /// Variation of Headings111 with shifted levels.
    /// </summary>
    Headings111Shifted,
    /// <summary>
    /// Lowercase letters followed by a bracket.
    /// </summary>
    LowerLetterWithBracket,
    /// <summary>
    /// Lowercase letters followed by a dot.
    /// </summary>
    LowerLetterWithDot,
    /// <summary>
    /// Uppercase letters followed by a dot.
    /// </summary>
    UpperLetterWithDot,
    /// <summary>
    /// Uppercase letters followed by a bracket.
    /// </summary>
    UpperLetterWithBracket,
    /// <summary>
    /// Custom numbering defined by the user.
    /// </summary>
    Custom
}
