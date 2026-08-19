namespace OfficeIMO.Html;

internal sealed class OfficeHtmlThemePalette {
    private OfficeHtmlThemePalette(
        string accent,
        string accentDark,
        string accentSoft,
        string pageBackground,
        string surface,
        string panel,
        string border,
        string borderStrong,
        string text,
        string heading,
        string muted,
        string tableHeader,
        string tableStripe,
        string codeBackground,
        string controlDisabled,
        string warning,
        string warningSoft,
        string fontBody,
        string fontHeading,
        string fontMono,
        string fontSize) {
        Accent = accent;
        AccentDark = accentDark;
        AccentSoft = accentSoft;
        PageBackground = pageBackground;
        Surface = surface;
        Panel = panel;
        Border = border;
        BorderStrong = borderStrong;
        Text = text;
        Heading = heading;
        Muted = muted;
        TableHeader = tableHeader;
        TableStripe = tableStripe;
        CodeBackground = codeBackground;
        ControlDisabled = controlDisabled;
        Warning = warning;
        WarningSoft = warningSoft;
        FontBody = fontBody;
        FontHeading = fontHeading;
        FontMono = fontMono;
        FontSize = fontSize;
    }

    internal string Accent { get; }
    internal string AccentDark { get; }
    internal string AccentSoft { get; }
    internal string PageBackground { get; }
    internal string Surface { get; }
    internal string Panel { get; }
    internal string Border { get; }
    internal string BorderStrong { get; }
    internal string Text { get; }
    internal string Heading { get; }
    internal string Muted { get; }
    internal string TableHeader { get; }
    internal string TableStripe { get; }
    internal string CodeBackground { get; }
    internal string ControlDisabled { get; }
    internal string Warning { get; }
    internal string WarningSoft { get; }
    internal string FontBody { get; }
    internal string FontHeading { get; }
    internal string FontMono { get; }
    internal string FontSize { get; }

    internal static OfficeHtmlThemePalette Create(OfficeVisualThemeKind theme) {
        switch (theme) {
            case OfficeVisualThemeKind.Plain:
                return Build("#475569", "#334155", "#F1F5F9", "#F8FAFC", "#FFFFFF", "#F8FAFC", "#E2E8F0", "#CBD5E1", "#111827", "#0F172A", "#475569", "#F1F5F9", "#FAFAFA", "#F8FAFC", "#F1F5F9", "14px");
            case OfficeVisualThemeKind.TechnicalDocument:
                return Build("#047857", "#065F46", "#ECFDF5", "#F1F5F9", "#FFFFFF", "#F8FAFC", "#D1D5DB", "#94A3B8", "#0F172A", "#064E3B", "#475569", "#D1FAE5", "#F8FAFC", "#F1F5F9", "#F1F5F9", "13.5px");
            case OfficeVisualThemeKind.GitHubLike:
                return Build("#0969DA", "#0550AE", "#DDF4FF", "#F6F8FA", "#FFFFFF", "#F6F8FA", "#D0D7DE", "#8C959F", "#1F2328", "#1F2328", "#656D76", "#F6F8FA", "#FBFCFD", "#F6F8FA", "#F6F8FA", "14px", "-apple-system,BlinkMacSystemFont,\"Segoe UI\",sans-serif", "-apple-system,BlinkMacSystemFont,\"Segoe UI\",sans-serif");
            case OfficeVisualThemeKind.Compact:
                return Build("#0F766E", "#115E59", "#F0FDFA", "#F1F5F9", "#FFFFFF", "#F8FAFC", "#CBD5E1", "#94A3B8", "#111827", "#134E4A", "#475569", "#CCFBF1", "#F8FAFC", "#F1F5F9", "#F1F5F9", "13px");
            case OfficeVisualThemeKind.Report:
                return Build("#1D4ED8", "#1E3A8A", "#EFF6FF", "#EEF2F7", "#FFFFFF", "#F8FAFC", "#CBD5E1", "#94A3B8", "#111827", "#172554", "#475569", "#DBEAFE", "#F8FAFC", "#F1F5F9", "#F1F5F9", "14px");
            default:
                return Build("#2563EB", "#1E40AF", "#EFF6FF", "#EEF2F7", "#FFFFFF", "#F8FAFC", "#D1D5DB", "#9CA3AF", "#111827", "#1F3763", "#4B5563", "#E7E6E6", "#FAFAFA", "#F3F4F6", "#F3F4F6", "14px", "Calibri,\"Segoe UI\",Arial,sans-serif", "Cambria,Georgia,\"Times New Roman\",serif");
        }
    }

    private static OfficeHtmlThemePalette Build(
        string accent,
        string accentDark,
        string accentSoft,
        string pageBackground,
        string surface,
        string panel,
        string border,
        string borderStrong,
        string text,
        string heading,
        string muted,
        string tableHeader,
        string tableStripe,
        string codeBackground,
        string controlDisabled,
        string fontSize,
        string fontBody = "\"Segoe UI\",Arial,sans-serif",
        string fontHeading = "\"Segoe UI\",Arial,sans-serif") =>
        new OfficeHtmlThemePalette(
            accent,
            accentDark,
            accentSoft,
            pageBackground,
            surface,
            panel,
            border,
            borderStrong,
            text,
            heading,
            muted,
            tableHeader,
            tableStripe,
            codeBackground,
            controlDisabled,
            "#B45309",
            "#FFFBEB",
            fontBody,
            fontHeading,
            "Consolas,\"SFMono-Regular\",\"Liberation Mono\",monospace",
            fontSize);
}
