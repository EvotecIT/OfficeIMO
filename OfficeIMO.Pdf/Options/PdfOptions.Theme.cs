namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private OfficeTheme? _theme;
    private PdfTheme? _themeDefaults;

    /// <summary>
    /// Cross-format theme that supplies composition defaults for headings, tables, lists, panels, rules and
    /// footer text when no explicit default style is configured. When null, <see cref="OfficeTheme.Default"/> is used.
    /// Explicit default styles and block styles always take precedence.
    /// </summary>
    public OfficeTheme? Theme {
        get => _theme;
        set {
            _theme = value;
            _themeDefaults = null;
            _themeDeferredTableStyle = null;
        }
    }

    internal OfficeTheme EffectiveTheme => _theme ?? OfficeTheme.Default;

    /// <summary>Footer text uses the theme's muted text colour unless a footer colour is set.</summary>
    internal PdfColor? EffectiveFooterTextColor => FooterTextColor ?? PdfColor.FromOfficeColor(EffectiveTheme.Palette.TextMuted);

    private PdfTableStyle? _themeDeferredTableStyle;

    /// <summary>
    /// Deferred tables stream their rows and cannot measure every row for column fitting, so the theme
    /// table style keeps its look there with columns sized by weight.
    /// </summary>
    internal PdfTableStyle? DefaultDeferredTableStyleSnapshot {
        get {
            if (HasExplicitDefaultTableStyleSnapshot) return DefaultTableStyleSnapshot;
            if (_themeDeferredTableStyle == null && ThemeDefaults.TableStyleSnapshot is PdfTableStyle themed) {
                PdfTableStyle copy = themed.Clone();
                copy.AutoFitColumns = false;
                _themeDeferredTableStyle = copy;
            }
            return _themeDeferredTableStyle;
        }
    }

    internal PdfTheme ThemeDefaults => _themeDefaults ??= PdfTheme.FromOfficeTheme(EffectiveTheme);
}
