namespace PSWriteOffice.Services;

/// <summary>Command vocabulary for native document restrictions; each format supports a subset.</summary>
public enum OfficeProtectionRestriction {
    /// <summary>Word read-only editing.</summary>
    ReadOnly,
    /// <summary>Word comment-only editing.</summary>
    Comments,
    /// <summary>Word tracked-change editing.</summary>
    TrackedChanges,
    /// <summary>Word form-field editing.</summary>
    Forms,
    /// <summary>Excel workbook structure protection.</summary>
    Structure,
    /// <summary>Excel workbook window protection.</summary>
    Windows,
    /// <summary>Excel read-only opening recommendation and write reservation.</summary>
    ReadOnlyRecommended,
    /// <summary>Withhold PDF printing, including high quality printing.</summary>
    Print,
    /// <summary>Withhold PDF content copying.</summary>
    Copy,
    /// <summary>Withhold PDF content modification.</summary>
    Modify,
    /// <summary>Withhold PDF annotation changes.</summary>
    Annotate,
    /// <summary>Withhold PDF form filling.</summary>
    FillForms,
    /// <summary>Withhold PDF document assembly.</summary>
    Assemble
}
