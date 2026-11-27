namespace OfficeIMO.Word;

/// <summary>
/// Specifies how a Word section begins.
/// </summary>
public enum WordSectionBreakType {
    /// <summary>Starts the section on the next page.</summary>
    NextPage,

    /// <summary>Starts the section in the next column.</summary>
    NextColumn,

    /// <summary>Starts the section without forcing a new page.</summary>
    Continuous,

    /// <summary>Starts the section on the next even-numbered page.</summary>
    EvenPage,

    /// <summary>Starts the section on the next odd-numbered page.</summary>
    OddPage
}

internal static class WordSectionBreakTypeExtensions {
    internal static string ToOpenXmlToken(this WordSectionBreakType breakType) => breakType switch {
        WordSectionBreakType.NextPage => "nextPage",
        WordSectionBreakType.NextColumn => "nextColumn",
        WordSectionBreakType.Continuous => "continuous",
        WordSectionBreakType.EvenPage => "evenPage",
        WordSectionBreakType.OddPage => "oddPage",
        _ => throw new ArgumentOutOfRangeException(nameof(breakType), breakType, "Unsupported Word section break type.")
    };

    internal static WordSectionBreakType FromOpenXmlToken(string? token) => token switch {
        null or "nextPage" => WordSectionBreakType.NextPage,
        "nextColumn" => WordSectionBreakType.NextColumn,
        "continuous" => WordSectionBreakType.Continuous,
        "evenPage" => WordSectionBreakType.EvenPage,
        "oddPage" => WordSectionBreakType.OddPage,
        _ => throw new FormatException("Unsupported Word section break type: " + token)
    };
}
