namespace OfficeIMO.Rtf;

/// <summary>
/// Page range to which page borders apply.
/// </summary>
public enum RtfPageBorderScope {
    /// <summary>Apply to all pages in the section.</summary>
    AllPagesInSection,

    /// <summary>Apply only to the first page in the section.</summary>
    FirstPageInSection,

    /// <summary>Apply to all pages except the first page in the section.</summary>
    AllExceptFirstPageInSection,

    /// <summary>Apply to the whole document.</summary>
    WholeDocument
}
