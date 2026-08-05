using DocumentFormat.OpenXml.Wordprocessing;

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
    internal static SectionMarkValues ToOpenXml(this WordSectionBreakType breakType) {
        return breakType switch {
            WordSectionBreakType.NextPage => SectionMarkValues.NextPage,
            WordSectionBreakType.NextColumn => SectionMarkValues.NextColumn,
            WordSectionBreakType.Continuous => SectionMarkValues.Continuous,
            WordSectionBreakType.EvenPage => SectionMarkValues.EvenPage,
            WordSectionBreakType.OddPage => SectionMarkValues.OddPage,
            _ => throw new ArgumentOutOfRangeException(nameof(breakType), breakType, "Unsupported Word section break type.")
        };
    }

    internal static WordSectionBreakType ToOfficeEnum(this SectionMarkValues breakType) {
        return breakType switch {
            _ when breakType == SectionMarkValues.NextPage => WordSectionBreakType.NextPage,
            _ when breakType == SectionMarkValues.NextColumn => WordSectionBreakType.NextColumn,
            _ when breakType == SectionMarkValues.Continuous => WordSectionBreakType.Continuous,
            _ when breakType == SectionMarkValues.EvenPage => WordSectionBreakType.EvenPage,
            _ when breakType == SectionMarkValues.OddPage => WordSectionBreakType.OddPage,
            _ => throw new ArgumentOutOfRangeException(nameof(breakType), breakType, "Unsupported Word section break type.")
        };
    }
}
