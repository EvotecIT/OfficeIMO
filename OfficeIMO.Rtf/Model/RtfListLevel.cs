namespace OfficeIMO.Rtf;

/// <summary>
/// Level inside an RTF list definition.
/// </summary>
public sealed class RtfListLevel {
    /// <summary>Creates a list level.</summary>
    public RtfListLevel(int levelIndex, RtfListKind kind = RtfListKind.Decimal) {
        if (levelIndex < 0) throw new ArgumentOutOfRangeException(nameof(levelIndex), "List level cannot be negative.");
        LevelIndex = levelIndex;
        Kind = kind;
    }

    /// <summary>Zero-based list level index.</summary>
    public int LevelIndex { get; }

    /// <summary>Basic marker kind.</summary>
    public RtfListKind Kind { get; set; }

    /// <summary>RTF numbering format code from <c>\levelnfc</c>, when present.</summary>
    public int? NumberFormat { get; set; }

    /// <summary>RTF numbering format code from <c>\levelnfcn</c>, when present.</summary>
    public int? NumberFormatN { get; set; }

    /// <summary>List marker alignment from <c>\leveljc</c>.</summary>
    public RtfListLevelAlignment? Alignment { get; set; }

    /// <summary>Bidirectional-aware list marker alignment from <c>\leveljcn</c>.</summary>
    public RtfListLevelAlignment? AlignmentN { get; set; }

    /// <summary>Character that follows the list marker.</summary>
    public RtfListLevelFollowCharacter? FollowCharacter { get; set; }

    /// <summary>Starting number for numbered lists.</summary>
    public int? StartAt { get; set; }

    /// <summary>Legacy minimum distance from the number's right edge to the paragraph text, in twips.</summary>
    public int? SpaceTwips { get; set; }

    /// <summary>Legacy minimum distance from left indent to paragraph text, in twips.</summary>
    public int? IndentTwips { get; set; }

    /// <summary>Whether previous-level numbers should be displayed as Arabic numbers.</summary>
    public bool? LegalNumbering { get; set; }

    /// <summary>Whether this level does not restart when a higher level increments.</summary>
    public bool? NoRestart { get; set; }

    /// <summary>Optional picture bullet index from <c>\levelpicture</c>.</summary>
    public int? PictureIndex { get; set; }

    /// <summary>Whether picture bullets should keep their original size.</summary>
    public bool PictureNoSize { get; set; }

    /// <summary>Marker text from <c>\leveltext</c>.</summary>
    public string? Text { get; set; }

    /// <summary>Marker number placeholders from <c>\levelnumbers</c>.</summary>
    public string? Numbers { get; set; }

    /// <summary>Left indentation in twips.</summary>
    public int? LeftIndentTwips { get; set; }

    /// <summary>First-line indentation in twips.</summary>
    public int? FirstLineIndentTwips { get; set; }
}
