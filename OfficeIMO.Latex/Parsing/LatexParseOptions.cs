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
    private readonly HashSet<string> _verbatimEnvironmentNames = new HashSet<string>(StringComparer.Ordinal) {
        "verbatim", "verbatim*", "Verbatim", "lstlisting", "minted", "comment"
    };

    /// <summary>Creates bounded parsing options for the requested named profile.</summary>
    public static LatexParseOptions CreateProfile(LatexDocumentProfile profile) =>
        profile switch {
            LatexDocumentProfile.OfficeIMO => new LatexParseOptions(),
            LatexDocumentProfile.PreserveOnly => new LatexParseOptions {
                Profile = LatexDocumentProfile.PreserveOnly
            },
            _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown LaTeX document profile.")
        };

    /// <summary>Semantic profile. Defaults to OfficeIMO.</summary>
    public LatexDocumentProfile Profile { get; set; } = LatexDocumentProfile.OfficeIMO;
    /// <summary>Maximum input characters.</summary>
    public int? MaximumInputLength { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum encoded bytes accepted by file and stream loading APIs.</summary>
    public long? MaximumInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum tokens.</summary>
    public int MaximumTokenCount { get; set; } = 2_000_000;
    /// <summary>Maximum nested groups and environments.</summary>
    public int MaximumNestingDepth { get; set; } = 128;

    /// <summary>
    /// Environment names whose bodies are opaque. Commands, comments, groups, math, and table separators
    /// inside these environments are retained as source but never interpreted semantically.
    /// </summary>
    public ISet<string> VerbatimEnvironmentNames => _verbatimEnvironmentNames;

    /// <summary>Macro expansion mode. Parsing itself never expands macros.</summary>
    public LatexMacroExpansion MacroExpansion { get; set; } = LatexMacroExpansion.None;

    /// <summary>Maximum recursive safe macro expansion depth.</summary>
    public int MaximumExpansionDepth { get; set; } = 16;

    /// <summary>Maximum characters produced by explicit safe macro expansion.</summary>
    public int MaximumExpansionLength { get; set; } = 16 * 1024 * 1024;

    /// <summary>Maximum input characters accepted by one explicit safe macro expansion step.</summary>
    public int MaximumExpansionInputLength { get; set; } = 64 * 1024 * 1024;

    /// <summary>Maximum tokens consumed across one explicit safe macro expansion.</summary>
    public int MaximumExpansionTokenCount { get; set; } = 2_000_000;

    internal void ValidateNamedModes() {
        if (Profile != LatexDocumentProfile.OfficeIMO && Profile != LatexDocumentProfile.PreserveOnly) {
            throw new ArgumentOutOfRangeException(nameof(Profile), Profile, "Unknown LaTeX document profile.");
        }
        if (MacroExpansion != LatexMacroExpansion.None
            && MacroExpansion != LatexMacroExpansion.SafeSimpleDefinitions) {
            throw new ArgumentOutOfRangeException(nameof(MacroExpansion), MacroExpansion,
                "Unknown LaTeX macro expansion mode.");
        }
    }
}
