namespace OfficeIMO.Rtf;

/// <summary>
/// Restart behavior for section line numbering.
/// </summary>
public enum RtfLineNumberRestart {
    /// <summary>Line numbers continue from the previous section.</summary>
    Continuous,

    /// <summary>Line numbers restart for each section.</summary>
    EachSection,

    /// <summary>Line numbers restart on each page.</summary>
    EachPage
}
