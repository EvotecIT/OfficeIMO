namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private OfficeTheme? _theme;
    private bool _useOfficeThemeDefaults = true;
    private PdfTheme? _themeDefaults;
    private PdfHeadingStyles? _themeHeadingFontSource;
    private PdfHeadingStyles? _themeHeadingsWithDocumentFont;
    private PdfTableStyle? _themeDeferredTableStyle;
    private bool _hasExplicitDefaultFontSize;
    private bool _hasExplicitFooterFontSize;
    private OfficeColor? _defaultTextColor;

    /// <summary>
    /// Supplies defaults for new PDF composition. Null selects OfficeTheme.Default.
    /// Explicit option styles, fonts and sizes, and block/run formatting take precedence.
    /// Loading a PDF and source-format layout imports do not apply this theme.
    /// </summary>
    public OfficeTheme? Theme {
        get => _theme;
        set {
            _theme = value;
            _themeDefaults = null;
            _themeDeferredTableStyle = null;
            MarkFontConfigurationAssigned();
            MarkDefaultFontSizeAssigned();
        }
    }

    internal OfficeTheme EffectiveTheme => _theme ?? OfficeTheme.Default;
    internal PdfTheme? ThemeDefaults => _useOfficeThemeDefaults
        ? _themeDefaults ??= PdfTheme.FromOfficeTheme(EffectiveTheme) : null;

    /// <summary>Resolves inherited heading typography without overriding an explicit document font.</summary>
    private PdfHeadingStyles? ThemeHeadingStylesSnapshot {
        get {
            PdfHeadingStyles? inherited = ThemeDefaults?.HeadingStylesSnapshot;
            if (!_hasExplicitDefaultFont || inherited == null) return inherited;
            if (!ReferenceEquals(_themeHeadingFontSource, inherited)) {
                PdfHeadingStyles withDocumentFont = inherited.Clone();
                for (int level = 1; level <= 3; level++) {
                    PdfHeadingStyle? style = withDocumentFont.GetSnapshot(level);
                    if (style == null) continue;
                    style.Font = null;
                    style.FontFamily = null;
                }
                _themeHeadingsWithDocumentFont = withDocumentFont;
                _themeHeadingFontSource = inherited;
            }
            return _themeHeadingsWithDocumentFont;
        }
    }

    /// <summary>Combines explicit per-level styles with inherited defaults resolved for this options scope.</summary>
    private PdfHeadingStyles? ResolveDefaultHeadingStylesSnapshot() {
        PdfHeadingStyles? inherited = ThemeHeadingStylesSnapshot;
        if (_defaultHeadingStyles == null) return inherited;
        if (inherited == null) return _defaultHeadingStyles;
        return new PdfHeadingStyles {
            Level1 = _defaultHeadingStyles.GetSnapshot(1) ?? inherited.GetSnapshot(1),
            Level2 = _defaultHeadingStyles.GetSnapshot(2) ?? inherited.GetSnapshot(2),
            Level3 = _defaultHeadingStyles.GetSnapshot(3) ?? inherited.GetSnapshot(3)
        };
    }

    internal OfficeColor? EffectiveFooterTextColor => FooterTextColor ?? (_useOfficeThemeDefaults
        ? PdfTheme.Resolve(EffectiveTheme.Footer.Color, EffectiveTheme) : (OfficeColor?)null);

    /// <summary>Creates options for source-backed PDF state and source-format layout.</summary>
    internal static PdfCreateOptions CreateNeutralOptions() => new PdfCreateOptions { _useOfficeThemeDefaults = false };

    /// <summary>Snapshots defaults without turning inherited values into explicit overrides.</summary>
    internal PdfCreateOptions CreateDocumentSnapshot(bool useThemeDefaults) {
        PdfCreateOptions snapshot = Clone();
        snapshot._useOfficeThemeDefaults = useThemeDefaults;
        snapshot._themeDefaults = null;
        snapshot._themeDeferredTableStyle = null;
        return snapshot;
    }

    /// <summary>Retains themed appearance with weighted columns on bounded deferred tables.</summary>
    internal PdfTableStyle? DefaultDeferredTableStyleSnapshot {
        get {
            if (!_useOfficeThemeDefaults || _hasExplicitDefaultTableStyle) return DefaultTableStyleSnapshot;
            if (_themeDeferredTableStyle == null && ThemeDefaults?.TableStyleSnapshot is PdfTableStyle themed) {
                _themeDeferredTableStyle = themed.Clone();
                _themeDeferredTableStyle.AutoFitColumns = false;
            }
            return _themeDeferredTableStyle;
        }
    }
}
