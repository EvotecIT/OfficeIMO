namespace OfficeIMO.Rtf;

/// <summary>
/// Section break behavior used by RTF section properties.
/// </summary>
public enum RtfSectionBreakKind {
    /// <summary>Start the next section on a new page.</summary>
    NextPage,

    /// <summary>Continue the next section without forcing a new page.</summary>
    Continuous,

    /// <summary>Start the next section in the next column.</summary>
    Column,

    /// <summary>Start the next section on the next even page.</summary>
    EvenPage,

    /// <summary>Start the next section on the next odd page.</summary>
    OddPage
}
