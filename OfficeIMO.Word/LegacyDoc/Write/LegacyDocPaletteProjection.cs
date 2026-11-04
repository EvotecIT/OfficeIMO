using OfficeIMO.Word.LegacyDoc.Model;
using System.Threading;

namespace OfficeIMO.Word.LegacyDoc.Write;

/// <summary>Collects palette loss for one synchronous native DOC encoding.</summary>
internal sealed class LegacyDocPaletteProjection : IDisposable {
    private static readonly AsyncLocal<LegacyDocPaletteProjection?> Active = new();
    private readonly LegacyDocPaletteProjection? _previous;
    private readonly Dictionary<string, OfficeDiagnostic> _diagnostics = new(StringComparer.Ordinal);

    internal LegacyDocPaletteProjection() { _previous = Active.Value; Active.Value = this; }

    // Validation admits representable colors without reporting unused conditional
    // styles. Only actual native paint emission records an approximation.
    internal static bool CanResolveColor(string? colorHex) =>
        LegacyDocColorPalette.TryGetIcoForHex(colorHex, out _) ||
        Active.Value != null && LegacyDocColorPalette.TryGetNearestIcoForHex(colorHex, out _);

    internal static bool TryResolveColor(string? colorHex, string location, out byte ico) {
        if (LegacyDocColorPalette.TryGetIcoForHex(colorHex, out ico)) return true;
        LegacyDocPaletteProjection? scope = Active.Value;
        if (scope == null || !LegacyDocColorPalette.TryGetNearestIcoForHex(colorHex, out ico)) return false;
        string source = colorHex!.Replace("#", string.Empty).ToUpperInvariant();
        string target = LegacyDocColorPalette.GetHexForIco(ico)!;
        string key = location + ":" + source;
        if (!scope._diagnostics.ContainsKey(key)) {
            scope._diagnostics.Add(key, new OfficeDiagnostic(OfficeDiagnosticCodes.WordDocColorApproximated,
                OfficeDiagnosticSeverity.Warning,
                $"Native DOC {location} color #{source} is approximated by palette color #{target}.",
                location, "Word", OfficeConversionLossKind.Approximation));
        }
        return true;
    }

    internal OfficeConversionReport CreateReport(WordDocument document, TimeSpan elapsed) =>
        new(document.FilePath ?? "Word", "Doc", itemsWritten: document.Sections.Count,
            elapsed: elapsed, diagnostics: _diagnostics.Values);

    public void Dispose() => Active.Value = _previous;
}
