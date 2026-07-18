using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared font-resolution diagnostics for dependency-free image exporters.</summary>
public static class OfficeImageExportFontDiagnostics {
    /// <summary>
    /// Reports when the renderer cannot use the first requested font family/style and must select
    /// a later family or the managed stroke fallback.
    /// </summary>
    public static OfficeImageExportDiagnostic? CreateSubstitutionDiagnostic(
        this OfficeFontFaceCollection fonts,
        string? text,
        string? familyNames,
        OfficeFontStyle style = OfficeFontStyle.Regular,
        string? source = null) {
        if (fonts == null) throw new ArgumentNullException(nameof(fonts));
        if (string.IsNullOrEmpty(text) || string.IsNullOrWhiteSpace(familyNames)) return null;

        List<string> families = ParseFamilies(familyNames!);
        if (families.Count == 0 || IsGenericFamily(families[0])) return null;

        OfficeFontStyle requestedStyle = OfficeFontFace.NormalizeStyle(style);
        for (int index = 0; index < families.Count; index++) {
            string family = families[index];
            OfficeTrueTypeFont? scoped = fonts.ResolveForText(text!, family, requestedStyle, out OfficeFontStyle resolvedStyle);
            if (scoped != null) {
                if (index == 0 && resolvedStyle == requestedStyle) return null;
                return CreateDiagnostic(
                    families[0],
                    requestedStyle,
                    family,
                    resolvedStyle,
                    scoped: true,
                    source);
            }

            OfficeTrueTypeFont? platform = OfficeTrueTypeFont.TryLoadFontFamily(family);
            if (platform != null && platform.HasGlyphs(text!)) {
                if (index == 0) return null;
                return CreateDiagnostic(
                    families[0],
                    requestedStyle,
                    family,
                    requestedStyle,
                    scoped: false,
                    source);
            }
        }

        return new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            OfficeImageExportDiagnosticCodes.FontSubstituted,
            "Font family '" + families[0] +
            "' could not be resolved with glyph coverage for this text; the dependency-free managed stroke fallback was used.",
            source,
            OfficeImageExportLossKind.Approximation);
    }

    /// <summary>Adds de-duplicated font diagnostics for all text in a drawing, including nested groups and patterns.</summary>
    public static void AppendFontDiagnostics(
        this OfficeDrawing drawing,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        string? source = null) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));

        var seen = new HashSet<string>(StringComparer.Ordinal);
        AppendDrawing(drawing, diagnostics, seen, source);
    }

    private static void AppendDrawing(
        OfficeDrawing drawing,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        HashSet<string> seen,
        string? source) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            switch (element) {
                case OfficeDrawingText text:
                    Append(
                        drawing.Fonts.CreateSubstitutionDiagnostic(
                            text.Text,
                            text.Font.FamilyName,
                            text.Font.Style,
                            source),
                        diagnostics,
                        seen);
                    break;
                case OfficeDrawingRichText richText:
                    foreach (OfficeRichTextRun run in richText.Runs) {
                        OfficeFontStyle style =
                            (run.Bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular) |
                            (run.Italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
                        Append(
                            drawing.Fonts.CreateSubstitutionDiagnostic(
                                run.Text,
                                run.FontFamily,
                                style,
                                source),
                            diagnostics,
                            seen);
                    }
                    break;
                case OfficeDrawingGroup group:
                    AppendDrawing(group.InnerDrawing, diagnostics, seen, source);
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    AppendDrawing(effectGroup.InnerDrawing, diagnostics, seen, source);
                    if (effectGroup.SoftMask != null) {
                        AppendDrawing(effectGroup.SoftMask.InnerDrawing, diagnostics, seen, source);
                    }
                    break;
                case OfficeDrawingTilingPattern pattern:
                    AppendDrawing(pattern.InnerTile, diagnostics, seen, source);
                    break;
            }
        }
    }

    private static void Append(
        OfficeImageExportDiagnostic? diagnostic,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        HashSet<string> seen) {
        if (diagnostic == null) return;
        string key = diagnostic.Code + "\n" + diagnostic.Message + "\n" + diagnostic.Source;
        if (seen.Add(key)) diagnostics.Add(diagnostic);
    }

    private static OfficeImageExportDiagnostic CreateDiagnostic(
        string requestedFamily,
        OfficeFontStyle requestedStyle,
        string resolvedFamily,
        OfficeFontStyle resolvedStyle,
        bool scoped,
        string? source) {
        string reason = string.Equals(requestedFamily, resolvedFamily, StringComparison.OrdinalIgnoreCase)
            ? "the requested " + DescribeStyle(requestedStyle) + " face was unavailable"
            : "the requested family was unavailable or lacked glyph coverage";
        string origin = scoped ? "caller-supplied" : "platform";
        return new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            OfficeImageExportDiagnosticCodes.FontSubstituted,
            "Font family '" + requestedFamily + "' was substituted with the " + origin +
            " '" + resolvedFamily + "' " + DescribeStyle(resolvedStyle) + " face because " + reason + ".",
            source,
            OfficeImageExportLossKind.Approximation);
    }

    private static List<string> ParseFamilies(string familyNames) {
        var families = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string rawFamily in familyNames.Split(',')) {
            string family = CleanFamilyName(rawFamily);
            if (family.Length > 0 && seen.Add(family)) families.Add(family);
        }
        return families;
    }

    private static string CleanFamilyName(string familyName) {
        string value = familyName.Trim();
        while (value.Length >= 2 &&
               ((value[0] == '"' && value[value.Length - 1] == '"') ||
                (value[0] == '\'' && value[value.Length - 1] == '\''))) {
            value = value.Substring(1, value.Length - 2).Trim();
        }
        return value;
    }

    private static bool IsGenericFamily(string family) {
        string normalized = family.Replace("-", string.Empty).Replace(" ", string.Empty);
        return string.Equals(normalized, "serif", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(normalized, "sans", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(normalized, "sansserif", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(normalized, "monospace", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(normalized, "mono", StringComparison.OrdinalIgnoreCase);
    }

    private static string DescribeStyle(OfficeFontStyle style) {
        OfficeFontStyle normalized = OfficeFontFace.NormalizeStyle(style);
        return normalized == OfficeFontStyle.Regular
            ? "regular"
            : normalized.ToString().Replace(", ", " ").ToLowerInvariant();
    }
}
