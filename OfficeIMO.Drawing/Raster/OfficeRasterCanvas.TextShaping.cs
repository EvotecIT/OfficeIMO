using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private const int MaxShapedTextCacheEntries = 4096;
    private Dictionary<ShapedTextKey, OfficeTrueTypeFont.ShapedTextRun?>? _shapedTextCache;
    private Dictionary<ShapedTextKey, OfficeManagedTextFallback>? _managedTextCache;

    private bool TryGetShapedTextRun(
        string text,
        OfficeTrueTypeFont font,
        out OfficeTrueTypeFont.ShapedTextRun shapedRun) {
        if (_textShapingProvider == null) {
            shapedRun = null!;
            return false;
        }

        _cancellationToken.ThrowIfCancellationRequested();
        var key = new ShapedTextKey(text, font);
        Dictionary<ShapedTextKey, OfficeTrueTypeFont.ShapedTextRun?> cache =
            _shapedTextCache ??= new Dictionary<ShapedTextKey, OfficeTrueTypeFont.ShapedTextRun?>();
        if (cache.TryGetValue(key, out OfficeTrueTypeFont.ShapedTextRun? cached)) {
            shapedRun = cached!;
            return cached != null;
        }

        string logicalText = OfficeArabicTextShaper.ToLogicalText(text);
        OfficeTextShapingResult? result = _textShapingProvider.ShapeText(new OfficeTextShapingRequest(
            logicalText,
            font.DisplayName ?? string.Empty,
            font.FontDataForShaping,
            isOpenTypeCff: false,
            font.UnitsPerEm,
            OfficeTextElements.ResolveBaseDirection(logicalText),
            _textShapingLanguage,
            _cancellationToken,
            font.CollectionIndex,
            cloneFontData: false));
        OfficeTrueTypeFont.ShapedTextRun? resolved =
            result == null ? null : font.CreateShapedTextRun(logicalText, result);
        if (cache.Count >= MaxShapedTextCacheEntries) cache.Clear();
        cache[key] = resolved;
        shapedRun = resolved!;
        return resolved != null;
    }

    private double MeasureResolvedText(string text, OfficeTrueTypeFont font, double fontSize) {
        if (TryGetShapedTextRun(text, font, out OfficeTrueTypeFont.ShapedTextRun run)) {
            return run.Measure(fontSize);
        }
        OfficeManagedTextFallback fallback = GetManagedTextFallback(text, font);
        return font.Measure(fallback.Text, fontSize);
    }

    private List<List<OfficePoint>> GetResolvedTextContours(
        string text,
        OfficeTrueTypeFont font,
        double x,
        double y,
        double fontSize) =>
        TryGetShapedTextRun(text, font, out OfficeTrueTypeFont.ShapedTextRun run)
            ? run.GetContours(x, y, fontSize)
            : font.GetTextContours(GetManagedTextFallback(text, font).Text, x, y, fontSize);

    private OfficeManagedTextFallback GetManagedTextFallback(string text, OfficeTrueTypeFont font) {
        _cancellationToken.ThrowIfCancellationRequested();
        var key = new ShapedTextKey(text, font);
        Dictionary<ShapedTextKey, OfficeManagedTextFallback> cache =
            _managedTextCache ??= new Dictionary<ShapedTextKey, OfficeManagedTextFallback>();
        if (cache.TryGetValue(key, out OfficeManagedTextFallback cached)) return cached;

        OfficeManagedTextFallback fallback = OfficeManagedTextShaper.Resolve(
            text,
            font,
            _cancellationToken);
        if (fallback.Used || fallback.Incomplete) ReportTextShapingFallback(fallback.Incomplete);
        if (cache.Count >= MaxShapedTextCacheEntries) cache.Clear();
        cache[key] = fallback;
        return fallback;
    }

    private void ReportTextShapingFallback(bool incomplete) {
        if (_diagnosticSink == null || HasTextShapingFallbackDiagnostic()) return;
        if (incomplete) {
            if (_reportedIncompleteTextShapingFallback) return;
            _reportedIncompleteTextShapingFallback = true;
            _diagnosticSink.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                OfficeImageExportDiagnosticCodes.TextShapingFallback,
                "Rendered complex text with a bounded fallback that cannot provide complete OpenType shaping or Unicode bidi behavior. Supply TextShapingProvider for premium script fidelity.",
                _diagnosticSource,
                OfficeImageExportLossKind.Approximation));
            return;
        }

        if (_reportedBoundedTextShapingFallback) return;
        _reportedBoundedTextShapingFallback = true;
        _diagnosticSink.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            OfficeImageExportDiagnosticCodes.TextShapingFallback,
            "Rendered complex text with the dependency-free core-Arabic and bidirectional fallback. Supply TextShapingProvider for full OpenType shaping.",
            _diagnosticSource,
            OfficeImageExportLossKind.Approximation));
    }

    private bool HasTextShapingFallbackDiagnostic() {
        if (_diagnosticSink == null) return false;
        foreach (OfficeImageExportDiagnostic diagnostic in _diagnosticSink) {
            if (diagnostic.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback &&
                string.Equals(diagnostic.Source, _diagnosticSource, StringComparison.Ordinal)) {
                return true;
            }
        }
        return false;
    }

    private readonly struct ShapedTextKey : IEquatable<ShapedTextKey> {
        internal ShapedTextKey(string text, OfficeTrueTypeFont font) {
            Text = text;
            Font = font;
        }

        private string Text { get; }
        private OfficeTrueTypeFont Font { get; }

        public bool Equals(ShapedTextKey other) =>
            ReferenceEquals(Font, other.Font) &&
            string.Equals(Text, other.Text, StringComparison.Ordinal);

        public override bool Equals(object? obj) =>
            obj is ShapedTextKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                return (StringComparer.Ordinal.GetHashCode(Text) * 397) ^
                       RuntimeHelpers.GetHashCode(Font);
            }
        }
    }
}
