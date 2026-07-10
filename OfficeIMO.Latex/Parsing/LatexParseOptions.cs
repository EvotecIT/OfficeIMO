namespace OfficeIMO.Latex;

/// <summary>Named bounded LaTeX document profile.</summary>
public enum LatexDocumentProfile {
    /// <summary>OfficeIMO LaTeX2e interoperability profile.</summary>
    OfficeIMO = 0,
    /// <summary>Only lossless structure; no profile-specific semantic assumptions.</summary>
    PreserveOnly
}

/// <summary>Opt-in macro expansion behavior.</summary>
public enum LatexMacroExpansion {
    /// <summary>Never expand macros.</summary>
    None = 0,
    /// <summary>Expand only document-local, structurally safe simple definitions under hard limits.</summary>
    SafeSimpleDefinitions
}

/// <summary>Options for dependency-free LaTeX parsing.</summary>
public sealed class LatexParseOptions {
    /// <summary>Semantic profile. Defaults to OfficeIMO.</summary>
    public LatexDocumentProfile Profile { get; set; } = LatexDocumentProfile.OfficeIMO;
    /// <summary>Maximum input characters.</summary>
    public int? MaximumInputLength { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum tokens.</summary>
    public int MaximumTokenCount { get; set; } = 2_000_000;
    /// <summary>Maximum nested groups and environments.</summary>
    public int MaximumNestingDepth { get; set; } = 128;

    /// <summary>Macro expansion mode. Parsing itself never expands macros.</summary>
    public LatexMacroExpansion MacroExpansion { get; set; } = LatexMacroExpansion.None;

    /// <summary>Maximum recursive safe macro expansion depth.</summary>
    public int MaximumExpansionDepth { get; set; } = 16;

    /// <summary>Maximum characters produced by explicit safe macro expansion.</summary>
    public int MaximumExpansionLength { get; set; } = 16 * 1024 * 1024;
}
