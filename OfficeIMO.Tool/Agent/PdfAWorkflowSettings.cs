using OfficeIMO.Pdf;

namespace OfficeIMO.Tool.Agent;

/// <summary>PDF/A conversion choices shared by the CLI and MCP adapters, named after <see cref="PdfAConversionOptions"/>.</summary>
internal sealed class PdfAWorkflowSettings {
    internal const int MaximumFonts = 64;
    private const long MaximumIccProfileBytes = 4L * 1024 * 1024;

    internal string Profile { get; init; } = "2b";
    internal string Attachments { get; init; } = "remove";
    internal string FontSubstitution { get; init; } = "installed-fonts";
    internal IReadOnlyList<string> Fonts { get; init; } = Array.Empty<string>();
    internal string? OutputIntent { get; init; }
    internal string? OutputConditionIdentifier { get; init; }

    /// <summary>Checks every choice before any file is read.</summary>
    internal void Validate() {
        _ = ParseProfile(Profile);
        _ = ParseAttachments(Attachments);
        _ = UsesInstalledFonts(FontSubstitution);
        if (Fonts.Count > MaximumFonts) throw new AgentUsageException("Supply at most " + MaximumFonts + " font faces.");
        _ = ParseFonts(Fonts);
        if (OutputIntent is null != OutputConditionIdentifier is null)
            throw new AgentUsageException("An output intent needs both its ICC profile and its output condition identifier.");
        if (OutputConditionIdentifier is { } identifier && (string.IsNullOrWhiteSpace(identifier) || identifier.Length > 256))
            throw new AgentUsageException("The output condition identifier must contain 1 through 256 characters.");
    }

    /// <summary>Every local file the conversion reads besides the source.</summary>
    internal IEnumerable<string> InputFiles() =>
        ParseFonts(Fonts).Select(font => font.Path).Concat(OutputIntent is null ? [] : [OutputIntent]);

    /// <summary>Builds library options from files already resolved through the host's path policy.</summary>
    internal PdfAConversionOptions ToOptions(Func<string, string> resolveInput) {
        Validate();
        var families = new List<PdfEmbeddedFontFamily>();
        foreach (IGrouping<string, FontFace> family in ParseFonts(Fonts).GroupBy(font => font.Family, StringComparer.OrdinalIgnoreCase)) {
            string? Face(string name) => family.SingleOrDefault(font => font.Face == name) is { } font ? resolveInput(font.Path) : null;
            string regular = Face("regular") ?? throw new AgentUsageException("Font family '" + family.Key + "' needs a regular face.");
            families.Add(PdfEmbeddedFontFamily.FromFiles(family.First().Family, regular, Face("bold"), Face("italic"), Face("bold-italic")));
        }
        bool installed = UsesInstalledFonts(FontSubstitution);
        PdfAFontSubstitution substitution = families.Count == 0
            ? installed ? PdfAFontSubstitution.InstalledFonts : PdfAFontSubstitution.None
            : installed ? PdfAFontSubstitution.Supplied(families.ToArray()).ThenInstalledFonts() : PdfAFontSubstitution.Supplied(families.ToArray());
        return new PdfAConversionOptions {
            Profile = ParseProfile(Profile),
            Attachments = ParseAttachments(Attachments),
            FontSubstitution = substitution,
            OutputIntent = OutputIntent is null ? null : new PdfOutputIntent(ReadIccProfile(resolveInput(OutputIntent)), OutputConditionIdentifier!)
        };
    }

    internal static PdfComplianceProfile ParseProfile(string value) => value.ToLowerInvariant() switch {
        "2b" => PdfComplianceProfile.PdfA2B,
        "2u" => PdfComplianceProfile.PdfA2U,
        "3b" => PdfComplianceProfile.PdfA3B,
        "3u" => PdfComplianceProfile.PdfA3U,
        "4" => PdfComplianceProfile.PdfA4,
        "4f" => PdfComplianceProfile.PdfA4F,
        _ => throw new AgentUsageException("profile must be 2b, 2u, 3b, 3u, 4, or 4f.")
    };

    private static PdfAAttachmentPolicy ParseAttachments(string value) => value.ToLowerInvariant() switch {
        "remove" => PdfAAttachmentPolicy.Remove,
        "refuse" => PdfAAttachmentPolicy.Refuse,
        _ => throw new AgentUsageException("attachments must be remove or refuse.")
    };

    private static bool UsesInstalledFonts(string value) => value.ToLowerInvariant() switch {
        "installed-fonts" => true,
        "none" => false,
        _ => throw new AgentUsageException("fontSubstitution must be installed-fonts or none.")
    };

    // family[:regular|bold|italic|bold-italic]=path, mirroring the faces of PdfEmbeddedFontFamily.FromFiles.
    private static List<FontFace> ParseFonts(IReadOnlyList<string> fonts) {
        var faces = new List<FontFace>(fonts.Count);
        foreach (string entry in fonts) {
            int split = entry.IndexOf('=');
            if (split <= 0 || split == entry.Length - 1) throw new AgentUsageException("Fonts use family=path or family:bold=path entries.");
            string name = entry[..split].Trim(), path = entry[(split + 1)..];
            int colon = name.LastIndexOf(':');
            string family = (colon < 0 ? name : name[..colon]).Trim(), face = colon < 0 ? "regular" : name[(colon + 1)..].Trim().ToLowerInvariant();
            if (family.Length is 0 or > 128) throw new AgentUsageException("Font family names must contain 1 through 128 characters.");
            if (face is not ("regular" or "bold" or "italic" or "bold-italic"))
                throw new AgentUsageException("Font faces are regular, bold, italic, or bold-italic.");
            if (faces.Any(existing => string.Equals(existing.Family, family, StringComparison.OrdinalIgnoreCase) && existing.Face == face))
                throw new AgentUsageException("Font family '" + family + "' lists its " + face + " face twice.");
            faces.Add(new FontFace(family, face, path));
        }
        return faces;
    }

    private static byte[] ReadIccProfile(string path) {
        using FileStream stream = File.OpenRead(path);
        if (stream.Length is <= 0 or > MaximumIccProfileBytes) throw new AgentUsageException("The ICC profile must contain 1 byte through 4 MiB.");
        byte[] profile = new byte[stream.Length];
        stream.ReadExactly(profile);
        return profile;
    }

    private sealed record FontFace(string Family, string Face, string Path);
}
