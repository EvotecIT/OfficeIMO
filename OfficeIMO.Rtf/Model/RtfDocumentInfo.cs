namespace OfficeIMO.Rtf;

/// <summary>
/// Document information stored in the RTF <c>info</c> destination.
/// </summary>
public sealed class RtfDocumentInfo {
    /// <summary>Generator metadata stored in the ignorable RTF <c>generator</c> destination.</summary>
    public string? Generator { get; set; }

    /// <summary>Document title.</summary>
    public string? Title { get; set; }

    /// <summary>Document subject.</summary>
    public string? Subject { get; set; }

    /// <summary>Document author.</summary>
    public string? Author { get; set; }

    /// <summary>Document manager.</summary>
    public string? Manager { get; set; }

    /// <summary>Document company.</summary>
    public string? Company { get; set; }

    /// <summary>Document operator.</summary>
    public string? Operator { get; set; }

    /// <summary>Document category.</summary>
    public string? Category { get; set; }

    /// <summary>Document keywords.</summary>
    public string? Keywords { get; set; }

    /// <summary>Document comments.</summary>
    public string? Comments { get; set; }

    /// <summary>Hyperlink base address.</summary>
    public string? HyperlinkBase { get; set; }

    /// <summary>Document creation timestamp from the RTF <c>creatim</c> destination.</summary>
    public DateTime? Created { get; set; }

    /// <summary>Document revision timestamp from the RTF <c>revtim</c> destination.</summary>
    public DateTime? Revised { get; set; }

    /// <summary>Document print timestamp from the RTF <c>printim</c> destination.</summary>
    public DateTime? Printed { get; set; }

    /// <summary>Document backup timestamp from the RTF <c>buptim</c> destination.</summary>
    public DateTime? BackedUp { get; set; }

    /// <summary>Total editing time in minutes.</summary>
    public int? EditingMinutes { get; set; }

    /// <summary>Page count stored in the RTF <c>nofpages</c> field.</summary>
    public int? NumberOfPages { get; set; }

    /// <summary>Word count stored in the RTF <c>nofwords</c> field.</summary>
    public int? NumberOfWords { get; set; }

    /// <summary>Character count stored in the RTF <c>nofchars</c> field.</summary>
    public int? NumberOfCharacters { get; set; }

    /// <summary>Character count including whitespace stored in the RTF <c>nofcharsws</c> field.</summary>
    public int? NumberOfCharactersWithSpaces { get; set; }

    /// <summary>Internal RTF version number stored in <c>vern</c>.</summary>
    public int? InternalVersion { get; set; }
}
