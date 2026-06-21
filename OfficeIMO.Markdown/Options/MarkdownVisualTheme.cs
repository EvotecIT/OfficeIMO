namespace OfficeIMO.Markdown;

/// <summary>
/// Shared visual profile used to keep Markdown exports visually consistent across HTML, PDF, and Word.
/// </summary>
public sealed class MarkdownVisualTheme {
    private MarkdownVisualPalette _palette = new MarkdownVisualPalette();
    private MarkdownTableVisualStyle _table = new MarkdownTableVisualStyle();

    /// <summary>Creates a custom visual theme.</summary>
    public MarkdownVisualTheme() {
    }

    private MarkdownVisualTheme(MarkdownVisualThemeKind kind, string name, HtmlStyle htmlStyle) {
        Kind = kind;
        Name = name;
        HtmlStyle = htmlStyle;
    }

    /// <summary>Theme kind used for diagnostics and renderer mapping.</summary>
    public MarkdownVisualThemeKind Kind { get; set; } = MarkdownVisualThemeKind.WordLike;

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

    /// <summary>Creates one of the built-in shared visual themes.</summary>
    public static MarkdownVisualTheme Create(MarkdownVisualThemeKind kind) {
        switch (kind) {
            case MarkdownVisualThemeKind.Plain:
                return Plain();
            case MarkdownVisualThemeKind.WordLike:
                return WordLike();
            case MarkdownVisualThemeKind.TechnicalDocument:
                return TechnicalDocument();
            case MarkdownVisualThemeKind.GitHubLike:
                return GitHubLike();
            case MarkdownVisualThemeKind.Compact:
                return Compact();
            case MarkdownVisualThemeKind.Report:
                return Report();
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported Markdown visual theme kind.");
        }
    }

    /// <summary>Tries to create a built-in shared visual theme from an API or front-matter name.</summary>
    public static bool TryCreate(string? name, out MarkdownVisualTheme? theme) {
        theme = null;
        string normalized = NormalizeName(name);
        if (normalized.Length == 0) {
            return false;
        }

        switch (normalized) {
            case "plain":
            case "none":
                theme = Plain();
                return true;
            case "word":
            case "wordlike":
            case "worddocument":
                theme = WordLike();
                return true;
            case "technical":
            case "technicaldocument":
            case "docs":
            case "documentation":
                theme = TechnicalDocument();
                return true;
            case "github":
            case "githublike":
                theme = GitHubLike();
                return true;
            case "compact":
            case "dense":
                theme = Compact();
                return true;
            case "report":
            case "businessreport":
                theme = Report();
                return true;
            default:
                return false;
        }
    }

    /// <summary>Plain profile with minimal visual opinions.</summary>
    public static MarkdownVisualTheme Plain() => new MarkdownVisualTheme(MarkdownVisualThemeKind.Plain, "Plain", HtmlStyle.Clean) {
        _palette = new MarkdownVisualPalette(),
        _table = new MarkdownTableVisualStyle { BorderWidth = 0.5, CellPaddingX = 5, CellPaddingY = 4 }
    };

    /// <summary>Neutral Word-like profile for general documents.</summary>
    public static MarkdownVisualTheme WordLike() => new MarkdownVisualTheme(MarkdownVisualThemeKind.WordLike, "WordLike", HtmlStyle.Word) {
        _palette = new MarkdownVisualPalette {
            Accent = MarkdownColor.Parse("#2563EB"),
            Heading = MarkdownColor.Parse("#111827"),
            Text = MarkdownColor.Parse("#1F2937"),
            MutedText = MarkdownColor.Parse("#64748B"),
            Surface = MarkdownColor.Parse("#F8FAFC"),
            Border = MarkdownColor.Parse("#CBD5E1"),
            CodeBackground = MarkdownColor.Parse("#F6F8FA"),
            TableHeaderBackground = MarkdownColor.Parse("#EFF6FF"),
            TableHeaderText = MarkdownColor.Parse("#0F172A"),
            TableStripeBackground = MarkdownColor.Parse("#F8FAFC")
        }
    };

    /// <summary>Polished profile for technical guides, specs, and READMEs.</summary>
    public static MarkdownVisualTheme TechnicalDocument() => new MarkdownVisualTheme(MarkdownVisualThemeKind.TechnicalDocument, "TechnicalDocument", HtmlStyle.GithubAuto) {
        _palette = new MarkdownVisualPalette {
            Accent = MarkdownColor.Parse("#0969DA"),
            Heading = MarkdownColor.Parse("#0F172A"),
            Text = MarkdownColor.Parse("#0F172A"),
            MutedText = MarkdownColor.Parse("#64748B"),
            Surface = MarkdownColor.Parse("#F8FAFC"),
            Border = MarkdownColor.Parse("#CBD5E1"),
            CodeBackground = MarkdownColor.Parse("#F6F8FA"),
            TableHeaderBackground = MarkdownColor.Parse("#E0F2FE"),
            TableHeaderText = MarkdownColor.Parse("#0F172A"),
            TableStripeBackground = MarkdownColor.Parse("#F8FAFC")
        },
        _table = new MarkdownTableVisualStyle { BorderWidth = 0.5, CellPaddingX = 6, CellPaddingY = 5 }
    };

    /// <summary>GitHub-inspired profile for README-style exports.</summary>
    public static MarkdownVisualTheme GitHubLike() => new MarkdownVisualTheme(MarkdownVisualThemeKind.GitHubLike, "GitHubLike", HtmlStyle.GithubAuto) {
        _palette = new MarkdownVisualPalette {
            Accent = MarkdownColor.Parse("#0969DA"),
            Heading = MarkdownColor.Parse("#24292F"),
            Text = MarkdownColor.Parse("#24292F"),
            MutedText = MarkdownColor.Parse("#57606A"),
            Surface = MarkdownColor.Parse("#F6F8FA"),
            Border = MarkdownColor.Parse("#D0D7DE"),
            CodeBackground = MarkdownColor.Parse("#F6F8FA"),
            TableHeaderBackground = MarkdownColor.Parse("#F6F8FA"),
            TableHeaderText = MarkdownColor.Parse("#24292F"),
            TableStripeBackground = MarkdownColor.Parse("#F6F8FA")
        }
    };

    /// <summary>Compact profile for dense technical notes.</summary>
    public static MarkdownVisualTheme Compact() => new MarkdownVisualTheme(MarkdownVisualThemeKind.Compact, "Compact", HtmlStyle.Clean) {
        _palette = new MarkdownVisualPalette {
            Accent = MarkdownColor.Parse("#2563EB"),
            Heading = MarkdownColor.Parse("#1F2937"),
            Text = MarkdownColor.Parse("#1F2937"),
            MutedText = MarkdownColor.Parse("#64748B"),
            Surface = MarkdownColor.Parse("#F8FAFC"),
            Border = MarkdownColor.Parse("#E2E8F0"),
            CodeBackground = MarkdownColor.Parse("#F8FAFC"),
            TableHeaderBackground = MarkdownColor.Parse("#F1F5F9"),
            TableHeaderText = MarkdownColor.Parse("#1F2937"),
            TableStripeBackground = MarkdownColor.Parse("#F8FAFC")
        },
        _table = new MarkdownTableVisualStyle { BorderWidth = 0.4, CellPaddingX = 4, CellPaddingY = 3 }
    };

    /// <summary>Report-oriented profile with stronger hierarchy and tables.</summary>
    public static MarkdownVisualTheme Report() => new MarkdownVisualTheme(MarkdownVisualThemeKind.Report, "Report", HtmlStyle.Word) {
        _palette = new MarkdownVisualPalette {
            Accent = MarkdownColor.Parse("#1E40AF"),
            Heading = MarkdownColor.Parse("#1E293B"),
            Text = MarkdownColor.Parse("#1E293B"),
            MutedText = MarkdownColor.Parse("#475569"),
            Surface = MarkdownColor.Parse("#EFF6FF"),
            Border = MarkdownColor.Parse("#BFDBFE"),
            CodeBackground = MarkdownColor.Parse("#F8FAFC"),
            TableHeaderBackground = MarkdownColor.Parse("#DBEAFE"),
            TableHeaderText = MarkdownColor.Parse("#172554"),
            TableStripeBackground = MarkdownColor.Parse("#EFF6FF")
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
        if (accent != null) palette.Accent = MarkdownColor.Parse(accent);
        if (heading != null) palette.Heading = MarkdownColor.Parse(heading);
        if (text != null) palette.Text = MarkdownColor.Parse(text);
        if (mutedText != null) palette.MutedText = MarkdownColor.Parse(mutedText);
        if (background != null) palette.Background = MarkdownColor.Parse(background);
        if (surface != null) palette.Surface = MarkdownColor.Parse(surface);
        if (border != null) palette.Border = MarkdownColor.Parse(border);
        if (codeBackground != null) palette.CodeBackground = MarkdownColor.Parse(codeBackground);
        if (tableHeaderBackground != null) palette.TableHeaderBackground = MarkdownColor.Parse(tableHeaderBackground);
        if (tableHeaderText != null) palette.TableHeaderText = MarkdownColor.Parse(tableHeaderText);
        if (tableStripeBackground != null) palette.TableStripeBackground = MarkdownColor.Parse(tableStripeBackground);
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
        _palette.Accent = MarkdownColor.Parse(accent);
        _palette.Heading = MarkdownColor.Parse(heading);
        _palette.Surface = MarkdownColor.Parse(surface);
        _palette.Border = MarkdownColor.Parse(border);
        _palette.TableHeaderBackground = MarkdownColor.Parse(tableHeader);
        _palette.TableStripeBackground = MarkdownColor.Parse(surface);
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
}
