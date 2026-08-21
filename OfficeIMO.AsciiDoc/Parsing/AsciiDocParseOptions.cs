namespace OfficeIMO.AsciiDoc;

/// <summary>Named bounded AsciiDoc document profile.</summary>
public enum AsciiDocDocumentProfile {
    /// <summary>OfficeIMO AsciiDoc interoperability profile with typed common constructs.</summary>
    OfficeIMO = 0,
    /// <summary>Lossless source preservation without enabling preprocessing or custom extensions.</summary>
    PreserveOnly
}

/// <summary>Options controlling lossless AsciiDoc parsing.</summary>
public sealed class AsciiDocParseOptions {
    /// <summary>Creates parsing options for the requested named profile.</summary>
    public static AsciiDocParseOptions CreateProfile(AsciiDocDocumentProfile profile) =>
        profile switch {
            AsciiDocDocumentProfile.OfficeIMO => CreateOfficeIMOProfile(),
            AsciiDocDocumentProfile.PreserveOnly => CreatePreserveOnlyProfile(),
            _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown AsciiDoc document profile.")
        };

    /// <summary>Creates the OfficeIMO interoperability profile.</summary>
    public static AsciiDocParseOptions CreateOfficeIMOProfile() => new AsciiDocParseOptions();

    /// <summary>Creates the source-preserving profile.</summary>
    public static AsciiDocParseOptions CreatePreserveOnlyProfile() => new AsciiDocParseOptions {
        Profile = AsciiDocDocumentProfile.PreserveOnly
    };

    /// <summary>Selected bounded document profile. Defaults to OfficeIMO.</summary>
    public AsciiDocDocumentProfile Profile { get; set; } = AsciiDocDocumentProfile.OfficeIMO;

    /// <summary>
    /// Maximum accepted UTF-16 source length. Defaults to 64 MiB of characters. Set to null to disable the limit.
    /// </summary>
    public int? MaximumInputLength { get; set; } = 64 * 1024 * 1024;

    /// <summary>
    /// Maximum number of top-level source blocks. Defaults to 1,000,000. Set to null to disable the limit.
    /// </summary>
    public int? MaximumBlockCount { get; set; } = 1_000_000;

    /// <summary>Maximum nested inline formatting depth. Defaults to 64.</summary>
    public int MaximumInlineNestingDepth { get; set; } = 64;

    /// <summary>Maximum inline nodes created for one document. Defaults to 1,000,000.</summary>
    public int MaximumInlineNodeCount { get; set; } = 1_000_000;
}
