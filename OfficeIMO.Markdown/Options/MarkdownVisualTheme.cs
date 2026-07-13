namespace OfficeIMO.Markdown;

/// <summary>
/// Shared visual profile used to keep Markdown exports visually consistent across HTML, PDF, and Word.
/// </summary>
public sealed class MarkdownVisualTheme {
    private static readonly IReadOnlyList<MarkdownVisualThemePreset> BuiltInPresets = Array.AsReadOnly(new[] {
        new MarkdownVisualThemePreset(
            OfficeVisualThemeKind.Plain,
            "Plain",
            "Minimal styling that leaves renderers close to their plain defaults.",
            HtmlStyle.Clean,
            "none"),
        new MarkdownVisualThemePreset(
            OfficeVisualThemeKind.WordLike,
            "WordLike",
            "Clean document styling for general Word, HTML, and PDF exports.",
            HtmlStyle.Word,
            "word", "word-document"),
        new MarkdownVisualThemePreset(
            OfficeVisualThemeKind.TechnicalDocument,
            "TechnicalDocument",
            "Polished technical-document styling for guides, READMEs, and specifications.",
            HtmlStyle.GithubAuto,
            "technical", "docs", "documentation"),
        new MarkdownVisualThemePreset(
            OfficeVisualThemeKind.GitHubLike,
            "GitHubLike",
            "GitHub-inspired Markdown styling for README-style exports.",
            HtmlStyle.GithubAuto,
            "github"),
        new MarkdownVisualThemePreset(
            OfficeVisualThemeKind.Compact,
            "Compact",
            "Dense document styling for technical notes and command output.",
            HtmlStyle.Clean,
            "dense"),
        new MarkdownVisualThemePreset(
            OfficeVisualThemeKind.Report,
            "Report",
            "Report-oriented styling with stronger tables and section hierarchy.",
            HtmlStyle.Word,
            "business-report")
    });

    private static readonly IReadOnlyList<MarkdownColorSchemeKind> BuiltInColorSchemes = Array.AsReadOnly(new[] {
        MarkdownColorSchemeKind.Default,
        MarkdownColorSchemeKind.Blue,
        MarkdownColorSchemeKind.Emerald,
        MarkdownColorSchemeKind.Indigo,
        MarkdownColorSchemeKind.Rose,
        MarkdownColorSchemeKind.Amber,
        MarkdownColorSchemeKind.Slate
    });

    private MarkdownVisualPalette _palette = new MarkdownVisualPalette();
    private MarkdownTableVisualStyle _table = new MarkdownTableVisualStyle();

    /// <summary>Creates a custom visual theme.</summary>
    public MarkdownVisualTheme() {
    }

    private MarkdownVisualTheme(OfficeVisualThemeKind kind, string name, HtmlStyle htmlStyle) {
        Kind = kind;
        Name = name;
        HtmlStyle = htmlStyle;
    }

    /// <summary>Theme kind used for diagnostics and renderer mapping.</summary>
    public OfficeVisualThemeKind Kind { get; set; } = OfficeVisualThemeKind.WordLike;

    /// <summary>Human-readable theme name.</summary>
    public string Name { get; set; } = "Custom";

    /// <summary>Preferred built-in HTML style preset when rendering this theme to HTML.</summary>
    public HtmlStyle HtmlStyle { get; set; } = HtmlStyle.Word;

    /// <summary>Semantic color palette shared by exporters.</summary>
    public MarkdownVisualPalette Palette {
        get => _palette;
        set => _palette = value?.Clone() ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Table styling shared by exporters.</summary>
    public MarkdownTableVisualStyle Table {
        get => _table;
        set => _table = value?.Clone() ?? throw new ArgumentNullException(nameof(value));
    }

    internal MarkdownVisualPalette PaletteSnapshot => _palette;

    internal MarkdownTableVisualStyle TableSnapshot => _table;

    /// <summary>Built-in visual theme presets callers can offer as stable choices.</summary>
    public static IReadOnlyList<MarkdownVisualThemePreset> Presets => BuiltInPresets;

    /// <summary>Built-in accent color schemes that can be applied to any preset.</summary>
    public static IReadOnlyList<MarkdownColorSchemeKind> ColorSchemes => BuiltInColorSchemes;

    /// <summary>Creates one of the built-in shared visual themes.</summary>
    public static MarkdownVisualTheme Create(OfficeVisualThemeKind kind) {
        switch (kind) {
            case OfficeVisualThemeKind.Plain:
                return Plain();
            case OfficeVisualThemeKind.WordLike:
                return WordLike();
            case OfficeVisualThemeKind.TechnicalDocument:
                return TechnicalDocument();
            case OfficeVisualThemeKind.GitHubLike:
                return GitHubLike();
            case OfficeVisualThemeKind.Compact:
                return Compact();
            case OfficeVisualThemeKind.Report:
                return Report();
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported Markdown visual theme kind.");
        }
    }

    /// <summary>Creates one of the built-in shared visual themes and applies a built-in color scheme.</summary>
    public static MarkdownVisualTheme Create(OfficeVisualThemeKind kind, MarkdownColorSchemeKind colorScheme) {
        MarkdownVisualTheme theme = Create(kind);
        return colorScheme == MarkdownColorSchemeKind.Default ? theme : theme.WithColorScheme(colorScheme);
    }

    /// <summary>Default shared visual theme used when exporters are asked to produce styled output without an explicit theme.</summary>
    public static MarkdownVisualTheme Default() => WordLike();

    /// <summary>Returns a copy of the requested theme, or the default shared theme when enabled.</summary>
    public static MarkdownVisualTheme? ResolveOrDefault(MarkdownVisualTheme? theme, bool applyDefaultTheme = true) =>
        theme?.Clone() ?? (applyDefaultTheme ? Default() : null);

    /// <summary>Tries to create a built-in shared visual theme from an API or front-matter name.</summary>
    public static bool TryCreate(string? name, out MarkdownVisualTheme? theme) =>
        TryCreate(name, MarkdownColorSchemeKind.Default, out theme);

    /// <summary>Tries to create a built-in shared visual theme from an API or front-matter name and applies a built-in color scheme.</summary>
    public static bool TryCreate(string? name, MarkdownColorSchemeKind colorScheme, out MarkdownVisualTheme? theme) {
        theme = null;
        if (!TryResolveThemeKind(name, out OfficeVisualThemeKind kind)) {
            return false;
        }

        theme = Create(kind, colorScheme);
        return true;
    }

    /// <summary>Plain profile with minimal visual opinions.</summary>
    public static MarkdownVisualTheme Plain() => new MarkdownVisualTheme(OfficeVisualThemeKind.Plain, "Plain", HtmlStyle.Clean) {
        _palette = new MarkdownVisualPalette(),
        _table = new MarkdownTableVisualStyle { BorderWidth = 0.5, CellPaddingX = 5, CellPaddingY = 4 }
    };

    /// <summary>Neutral Word-like profile for general documents.</summary>
    public static MarkdownVisualTheme WordLike() => new MarkdownVisualTheme(OfficeVisualThemeKind.WordLike, "WordLike", HtmlStyle.Word) {
        _palette = new MarkdownVisualPalette {
            Accent = OfficeColor.Parse("#2563EB"),
            Heading = OfficeColor.Parse("#111827"),
            Text = OfficeColor.Parse("#1F2937"),
            MutedText = OfficeColor.Parse("#64748B"),
            Surface = OfficeColor.Parse("#F8FAFC"),
            Border = OfficeColor.Parse("#CBD5E1"),
            CodeBackground = OfficeColor.Parse("#F6F8FA"),
            TableHeaderBackground = OfficeColor.Parse("#EFF6FF"),
            TableHeaderText = OfficeColor.Parse("#0F172A"),
            TableStripeBackground = OfficeColor.Parse("#F8FAFC")
        }
    };

    /// <summary>Polished profile for technical guides, specs, and READMEs.</summary>
    public static MarkdownVisualTheme TechnicalDocument() => new MarkdownVisualTheme(OfficeVisualThemeKind.TechnicalDocument, "TechnicalDocument", HtmlStyle.GithubAuto) {
        _palette = new MarkdownVisualPalette {
            Accent = OfficeColor.Parse("#0969DA"),
            Heading = OfficeColor.Parse("#0F172A"),
            Text = OfficeColor.Parse("#0F172A"),
            MutedText = OfficeColor.Parse("#64748B"),
            Surface = OfficeColor.Parse("#F8FAFC"),
            Border = OfficeColor.Parse("#CBD5E1"),
            CodeBackground = OfficeColor.Parse("#F6F8FA"),
            TableHeaderBackground = OfficeColor.Parse("#E0F2FE"),
            TableHeaderText = OfficeColor.Parse("#0F172A"),
            TableStripeBackground = OfficeColor.Parse("#F8FAFC")
        },
        _table = new MarkdownTableVisualStyle { BorderWidth = 0.5, CellPaddingX = 6, CellPaddingY = 5 }
    };

    /// <summary>GitHub-inspired profile for README-style exports.</summary>
    public static MarkdownVisualTheme GitHubLike() => new MarkdownVisualTheme(OfficeVisualThemeKind.GitHubLike, "GitHubLike", HtmlStyle.GithubAuto) {
        _palette = new MarkdownVisualPalette {
            Accent = OfficeColor.Parse("#0969DA"),
            Heading = OfficeColor.Parse("#24292F"),
            Text = OfficeColor.Parse("#24292F"),
            MutedText = OfficeColor.Parse("#57606A"),
            Surface = OfficeColor.Parse("#F6F8FA"),
            Border = OfficeColor.Parse("#D0D7DE"),
            CodeBackground = OfficeColor.Parse("#F6F8FA"),
            TableHeaderBackground = OfficeColor.Parse("#F6F8FA"),
            TableHeaderText = OfficeColor.Parse("#24292F"),
            TableStripeBackground = OfficeColor.Parse("#F6F8FA")
        }
    };

    /// <summary>Compact profile for dense technical notes.</summary>
    public static MarkdownVisualTheme Compact() => new MarkdownVisualTheme(OfficeVisualThemeKind.Compact, "Compact", HtmlStyle.Clean) {
        _palette = new MarkdownVisualPalette {
            Accent = OfficeColor.Parse("#2563EB"),
            Heading = OfficeColor.Parse("#1F2937"),
            Text = OfficeColor.Parse("#1F2937"),
            MutedText = OfficeColor.Parse("#64748B"),
            Surface = OfficeColor.Parse("#F8FAFC"),
            Border = OfficeColor.Parse("#E2E8F0"),
            CodeBackground = OfficeColor.Parse("#F8FAFC"),
            TableHeaderBackground = OfficeColor.Parse("#F1F5F9"),
            TableHeaderText = OfficeColor.Parse("#1F2937"),
            TableStripeBackground = OfficeColor.Parse("#F8FAFC")
        },
        _table = new MarkdownTableVisualStyle { BorderWidth = 0.4, CellPaddingX = 4, CellPaddingY = 3 }
    };

    /// <summary>Report-oriented profile with stronger hierarchy and tables.</summary>
    public static MarkdownVisualTheme Report() => new MarkdownVisualTheme(OfficeVisualThemeKind.Report, "Report", HtmlStyle.Word) {
        _palette = new MarkdownVisualPalette {
            Accent = OfficeColor.Parse("#1E40AF"),
            Heading = OfficeColor.Parse("#1E293B"),
            Text = OfficeColor.Parse("#1E293B"),
            MutedText = OfficeColor.Parse("#475569"),
            Surface = OfficeColor.Parse("#EFF6FF"),
            Border = OfficeColor.Parse("#BFDBFE"),
            CodeBackground = OfficeColor.Parse("#F8FAFC"),
            TableHeaderBackground = OfficeColor.Parse("#DBEAFE"),
            TableHeaderText = OfficeColor.Parse("#172554"),
            TableStripeBackground = OfficeColor.Parse("#EFF6FF")
        },
        _table = new MarkdownTableVisualStyle { BorderWidth = 0.7, CellPaddingX = 6, CellPaddingY = 5 }
    };

    /// <summary>Returns a copy of this theme with the requested color scheme applied.</summary>
    public MarkdownVisualTheme WithColorScheme(MarkdownColorSchemeKind scheme) {
        MarkdownVisualTheme clone = Clone();
        clone.ApplyColorScheme(scheme);
        return clone;
    }

    /// <summary>Returns a copy of this theme with selected palette colors overridden.</summary>
    public MarkdownVisualTheme WithColors(
        string? accent = null,
        string? heading = null,
        string? text = null,
        string? mutedText = null,
        string? background = null,
        string? surface = null,
        string? border = null,
        string? codeBackground = null,
        string? tableHeaderBackground = null,
        string? tableHeaderText = null,
        string? tableStripeBackground = null) {
        MarkdownVisualTheme clone = Clone();
        MarkdownVisualPalette palette = clone._palette;
        if (accent != null) palette.Accent = OfficeColor.Parse(accent);
        if (heading != null) palette.Heading = OfficeColor.Parse(heading);
        if (text != null) palette.Text = OfficeColor.Parse(text);
        if (mutedText != null) palette.MutedText = OfficeColor.Parse(mutedText);
        if (background != null) palette.Background = OfficeColor.Parse(background);
        if (surface != null) palette.Surface = OfficeColor.Parse(surface);
        if (border != null) palette.Border = OfficeColor.Parse(border);
        if (codeBackground != null) palette.CodeBackground = OfficeColor.Parse(codeBackground);
        if (tableHeaderBackground != null) palette.TableHeaderBackground = OfficeColor.Parse(tableHeaderBackground);
        if (tableHeaderText != null) palette.TableHeaderText = OfficeColor.Parse(tableHeaderText);
        if (tableStripeBackground != null) palette.TableStripeBackground = OfficeColor.Parse(tableStripeBackground);
        return clone;
    }

    /// <summary>Returns a copy of this theme with table styling configured.</summary>
    public MarkdownVisualTheme WithTable(Action<MarkdownTableVisualStyle> configure) {
        if (configure == null) {
            throw new ArgumentNullException(nameof(configure));
        }

        MarkdownVisualTheme clone = Clone();
        configure(clone._table);
        return clone;
    }

    /// <summary>Creates a copy of this visual theme.</summary>
    public MarkdownVisualTheme Clone() => new MarkdownVisualTheme {
        Kind = Kind,
        Name = Name,
        HtmlStyle = HtmlStyle,
        _palette = _palette.Clone(),
        _table = _table.Clone()
    };

    private void ApplyColorScheme(MarkdownColorSchemeKind scheme) {
        switch (scheme) {
            case MarkdownColorSchemeKind.Default:
                return;
            case MarkdownColorSchemeKind.Blue:
                ApplyScheme("#2563EB", "#1E3A8A", "#EFF6FF", "#BFDBFE", "#DBEAFE");
                break;
            case MarkdownColorSchemeKind.Emerald:
                ApplyScheme("#059669", "#064E3B", "#ECFDF5", "#A7F3D0", "#D1FAE5");
                break;
            case MarkdownColorSchemeKind.Indigo:
                ApplyScheme("#4F46E5", "#312E81", "#EEF2FF", "#C7D2FE", "#E0E7FF");
                break;
            case MarkdownColorSchemeKind.Rose:
                ApplyScheme("#E11D48", "#881337", "#FFF1F2", "#FECDD3", "#FFE4E6");
                break;
            case MarkdownColorSchemeKind.Amber:
                ApplyScheme("#D97706", "#78350F", "#FFFBEB", "#FDE68A", "#FEF3C7");
                break;
            case MarkdownColorSchemeKind.Slate:
                ApplyScheme("#475569", "#0F172A", "#F8FAFC", "#CBD5E1", "#E2E8F0");
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(scheme), scheme, "Unsupported Markdown color scheme.");
        }
    }

    private void ApplyScheme(string accent, string heading, string surface, string border, string tableHeader) {
        _palette.Accent = OfficeColor.Parse(accent);
        _palette.Heading = OfficeColor.Parse(heading);
        _palette.Surface = OfficeColor.Parse(surface);
        _palette.Border = OfficeColor.Parse(border);
        _palette.TableHeaderBackground = OfficeColor.Parse(tableHeader);
        _palette.TableStripeBackground = OfficeColor.Parse(surface);
    }

    private static string NormalizeName(string? name) {
        if (string.IsNullOrWhiteSpace(name)) {
            return string.Empty;
        }

        string value = name!;
        var builder = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (char.IsLetterOrDigit(c)) {
                builder.Append(char.ToLowerInvariant(c));
            }
        }

        return builder.ToString();
    }

    private static bool TryResolveThemeKind(string? name, out OfficeVisualThemeKind kind) {
        kind = default;
        string normalized = NormalizeName(name);
        if (normalized.Length == 0) {
            return false;
        }

        foreach (MarkdownVisualThemePreset preset in BuiltInPresets) {
            if (NormalizeName(preset.Name) == normalized) {
                kind = preset.Kind;
                return true;
            }

            foreach (string alias in preset.Aliases) {
                if (NormalizeName(alias) == normalized) {
                    kind = preset.Kind;
                    return true;
                }
            }
        }

        return false;
    }
}
