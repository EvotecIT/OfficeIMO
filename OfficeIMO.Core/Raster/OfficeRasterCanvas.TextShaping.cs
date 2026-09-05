using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private const int MaxShapedTextCacheEntries = 4096;
    private const int MaximumTextOutlinePointsPerRun = 1_000_000;
    private Dictionary<ShapedTextKey, OfficeTextShapingResult?>? _shapedTextCache;
    private Dictionary<ShapedTextKey, OfficeManagedTextFallback>? _managedTextCache;
    private readonly OfficeCffOperationBudget _cffOperationBudget = new OfficeCffOperationBudget();

    private bool TryGetShapedTextRun(
        string text,
        IOfficeFontProgram font,
        OfficeTextFeatureSettings? featureSettings,
        out OfficeTextShapingResult shapedRun) {
        OfficeTextFeatureSettings resolvedFeatures = featureSettings ?? OfficeTextFeatureSettings.Default;
        IOfficeTextShapingProvider? provider = _textShapingProvider;
        if (provider == null && !resolvedFeatures.IsDefault) provider = OfficeManagedTextShapingProvider.Instance;
        if (provider == null) {
            shapedRun = null!;
            return false;
        }

        _cancellationToken.ThrowIfCancellationRequested();
        var key = new ShapedTextKey(text, font, resolvedFeatures);
        Dictionary<ShapedTextKey, OfficeTextShapingResult?> cache =
            _shapedTextCache ??= new Dictionary<ShapedTextKey, OfficeTextShapingResult?>();
        if (cache.TryGetValue(key, out OfficeTextShapingResult? cached)) {
            shapedRun = cached!;
            return cached != null;
        }

        string logicalText = OfficeArabicTextShaper.ToLogicalText(text);
        OfficeTextShapingResult? result = provider.ShapeText(new OfficeTextShapingRequest(
            logicalText,
            font.DisplayName ?? string.Empty,
            font.GetFontDataForShaping(),
            font.IsOpenTypeCff,
            font.UnitsPerEm,
            OfficeTextElements.ResolveBaseDirection(logicalText),
            _textShapingLanguage,
            _cancellationToken,
            font.CollectionIndex,
            (font as IOfficeVariableFontProgram)?.VariationCoordinatesForShaping,
            cloneFontData: false,
            fontProgramCacheKey: font,
            featureSettings: resolvedFeatures));
        OfficeTextShapingResult? resolved = result;
        if (cache.Count >= MaxShapedTextCacheEntries) cache.Clear();
        cache[key] = resolved;
        shapedRun = resolved!;
        return resolved != null;
    }

    private double MeasureResolvedText(string text, IOfficeFontProgram font, double fontSize, OfficeTextFeatureSettings? featureSettings = null) {
        if (font.ProvidesComplexTextLayout) return font.Measure(text, fontSize);
        if (TryGetShapedTextRun(text, font, featureSettings, out OfficeTextShapingResult run)) {
            return font.MeasureShapedText(OfficeArabicTextShaper.ToLogicalText(text), run, fontSize);
        }
        OfficeManagedTextFallback fallback = GetManagedTextFallback(text, font);
        return font.Measure(fallback.Text, fontSize);
    }

    private List<List<OfficePoint>> GetResolvedTextContours(
        string text,
        IOfficeFontProgram font,
        double x,
        double y,
        double fontSize,
        OfficeTextFeatureSettings? featureSettings = null) {
        if (font.ProvidesComplexTextLayout) {
            return GetBoundedTextContours(font, text, x, y, fontSize);
        }
        if (TryGetShapedTextRun(text, font, featureSettings, out OfficeTextShapingResult run)) {
            string logicalText = OfficeArabicTextShaper.ToLogicalText(text);
            if (font is IOfficeCffBoundedFontProgram cff) {
                return cff.GetShapedTextContoursBounded(
                    logicalText,
                    run,
                    x,
                    y,
                    fontSize,
                    MaximumTextOutlinePointsPerRun,
                    _cancellationToken,
                    _cffOperationBudget);
            }
            if (font is IOfficeBoundedFontProgram bounded) {
                return bounded.GetShapedTextContoursBounded(
                    logicalText,
                    run,
                    x,
                    y,
                    fontSize,
                    MaximumTextOutlinePointsPerRun,
                    _cancellationToken);
            }
            _cancellationToken.ThrowIfCancellationRequested();
            List<List<OfficePoint>> contours = font.GetShapedTextContours(logicalText, run, x, y, fontSize);
            _cancellationToken.ThrowIfCancellationRequested();
            EnsureBoundedContourPoints(contours, MaximumTextOutlinePointsPerRun);
            return contours;
        }
        return GetBoundedTextContours(
            font,
            GetManagedTextFallback(text, font).Text,
            x,
            y,
            fontSize);
    }

    private List<List<OfficePoint>> GetBoundedTextContours(
        IOfficeFontProgram font,
        string text,
        double x,
        double y,
        double fontSize) {
        if (font is IOfficeCffBoundedFontProgram cff) {
            return cff.GetTextContoursBounded(
                text,
                x,
                y,
                fontSize,
                MaximumTextOutlinePointsPerRun,
                _cancellationToken,
                _cffOperationBudget);
        }
        if (font is IOfficeBoundedFontProgram bounded) {
            return bounded.GetTextContoursBounded(
                text,
                x,
                y,
                fontSize,
                MaximumTextOutlinePointsPerRun,
                _cancellationToken);
        }
        _cancellationToken.ThrowIfCancellationRequested();
        List<List<OfficePoint>> contours = font.GetTextContours(text, x, y, fontSize);
        _cancellationToken.ThrowIfCancellationRequested();
        EnsureBoundedContourPoints(contours, MaximumTextOutlinePointsPerRun);
        return contours;
    }

    private static void EnsureBoundedContourPoints(
        IEnumerable<List<OfficePoint>> contours,
        int maximumPointCount) {
        int pointCount = 0;
        foreach (List<OfficePoint> contour in contours) {
            if (contour.Count > maximumPointCount - pointCount) {
                throw new InvalidOperationException("Font outline expansion exceeded the configured point budget.");
            }
            pointCount += contour.Count;
        }
    }

    private OfficeManagedTextFallback GetManagedTextFallback(string text, IOfficeFontProgram font) {
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
                OfficeConversionLossKind.Approximation));
            return;
        }

        if (_reportedBoundedTextShapingFallback) return;
        _reportedBoundedTextShapingFallback = true;
        _diagnosticSink.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            OfficeImageExportDiagnosticCodes.TextShapingFallback,
            "Rendered complex text with the dependency-free core-Arabic and bidirectional fallback. Supply TextShapingProvider for full OpenType shaping.",
            _diagnosticSource,
            OfficeConversionLossKind.Approximation));
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
        internal ShapedTextKey(string text, IOfficeFontProgram font, OfficeTextFeatureSettings? featureSettings = null) {
            Text = text;
            Font = font;
            FeatureSettings = featureSettings ?? OfficeTextFeatureSettings.Default;
        }

        private string Text { get; }
        private IOfficeFontProgram Font { get; }
        private OfficeTextFeatureSettings FeatureSettings { get; }

        public bool Equals(ShapedTextKey other) =>
            ReferenceEquals(Font, other.Font) &&
            FeatureSettings.Equals(other.FeatureSettings) &&
            string.Equals(Text, other.Text, StringComparison.Ordinal);

        public override bool Equals(object? obj) =>
            obj is ShapedTextKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                return (StringComparer.Ordinal.GetHashCode(Text) * 397) ^
                       RuntimeHelpers.GetHashCode(Font) ^ FeatureSettings.GetHashCode();
            }
        }
    }
}
