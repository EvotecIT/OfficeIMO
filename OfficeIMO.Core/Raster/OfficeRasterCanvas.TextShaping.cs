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
        out OfficeTextShapingResult shapedRun) => TryGetShapedTextRun(
            text,
            font,
            featureSettings,
            OfficeTextElements.ResolveBaseDirection(OfficeArabicTextShaper.ToLogicalText(text)),
            out shapedRun);

    private bool TryGetShapedTextRun(
        string text,
        IOfficeFontProgram font,
        OfficeTextFeatureSettings? featureSettings,
        OfficeTextDirection direction,
        out OfficeTextShapingResult shapedRun) {
        OfficeTextFeatureSettings resolvedFeatures = featureSettings ?? OfficeTextFeatureSettings.Default;
        IOfficeTextShapingProvider? provider = _textShapingProvider;
        if (provider == null && !resolvedFeatures.IsDefault) provider = OfficeManagedTextShapingProvider.Instance;
        if (provider == null) {
            shapedRun = null!;
            return false;
        }

        _cancellationToken.ThrowIfCancellationRequested();
        var key = new ShapedTextKey(text, font, resolvedFeatures, direction);
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
            direction,
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

    internal bool TryDrawVerticalText(
        string text,
        double x,
        double y,
        double width,
        double height,
        OfficeColor color,
        double fontSize,
        OfficeFontStyle style,
        string? fontFamily,
        OfficeTextFeatureSettings? featureSettings,
        string? fontPalette) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || width <= 0D || height <= 0D) return true;
        IOfficeFontProgram? font = ResolveTextFont(text, fontFamily, style, out OfficeFontStyle resolvedStyle);
        if (font == null ||
            !TryGetShapedTextRun(text, font, featureSettings, OfficeTextDirection.TopToBottom, out OfficeTextShapingResult run) ||
            !HasUsableVerticalPositioning(run)) {
            ReportTextShapingFallback(incomplete: true);
            return false;
        }

        double size = Math.Max(1D, fontSize);
        double originX = x + width / 2D;
        double originY = y;
        OfficeFontStyle simulatedStyle = style & ~resolvedStyle;
        if (font is OfficeTrueTypeFont trueType && trueType.TryGetShapedColorTextContours(
            text, run, originX, originY, size, fontPalette, color, MaximumTextOutlinePointsPerRun,
            _cancellationToken, out List<OfficeColorGlyphContours> colorLayers)) {
            AlignVerticalColorContoursToTop(colorLayers, y);
            foreach (OfficeColorGlyphContours layer in colorLayers) {
                if ((simulatedStyle & OfficeFontStyle.Italic) == OfficeFontStyle.Italic) SlantContours(layer.Contours, originY, size);
                FillContours(layer.Contours, layer.Color, OfficeFillRule.NonZero);
                if ((simulatedStyle & OfficeFontStyle.Bold) == OfficeFontStyle.Bold) {
                    OffsetContours(layer.Contours, size / 24D, 0D);
                    FillContours(layer.Contours, layer.Color, OfficeFillRule.NonZero);
                }
            }
            return true;
        }

        List<List<OfficePoint>> contours;
        if (font is IOfficeCffBoundedFontProgram cff) {
            contours = cff.GetShapedTextContoursBounded(text, run, originX, originY, size, MaximumTextOutlinePointsPerRun, _cancellationToken, _cffOperationBudget);
        } else if (font is IOfficeBoundedFontProgram bounded) {
            contours = bounded.GetShapedTextContoursBounded(text, run, originX, originY, size, MaximumTextOutlinePointsPerRun, _cancellationToken);
        } else {
            contours = font.GetShapedTextContours(text, run, originX, originY, size);
            EnsureBoundedContourPoints(contours, MaximumTextOutlinePointsPerRun);
        }
        AlignVerticalContoursToTop(contours, y);
        if ((simulatedStyle & OfficeFontStyle.Italic) == OfficeFontStyle.Italic) SlantContours(contours, originY, size);
        FillContours(contours, color, OfficeFillRule.NonZero);
        if ((simulatedStyle & OfficeFontStyle.Bold) == OfficeFontStyle.Bold) {
            OffsetContours(contours, size / 24D, 0D);
            FillContours(contours, color, OfficeFillRule.NonZero);
        }
        return true;
    }

    private static void AlignVerticalColorContoursToTop(List<OfficeColorGlyphContours> layers, double top) {
        double minimumY = double.PositiveInfinity;
        foreach (OfficeColorGlyphContours layer in layers) {
            minimumY = Math.Min(minimumY, FindMinimumContourY(layer.Contours));
        }
        if (double.IsPositiveInfinity(minimumY)) return;
        double offsetY = top - minimumY;
        foreach (OfficeColorGlyphContours layer in layers) OffsetContours(layer.Contours, 0D, offsetY);
    }

    private static void AlignVerticalContoursToTop(List<List<OfficePoint>> contours, double top) {
        // HarfBuzz vertical origins and synthetic providers use different baseline conventions.
        // Normalize actual ink before the caller clips the run to its declared text box.
        double minimumY = FindMinimumContourY(contours);
        if (!double.IsPositiveInfinity(minimumY)) OffsetContours(contours, 0D, top - minimumY);
    }

    private static double FindMinimumContourY(List<List<OfficePoint>> contours) {
        double minimumY = double.PositiveInfinity;
        foreach (List<OfficePoint> contour in contours) {
            foreach (OfficePoint point in contour) {
                if (!double.IsNaN(point.Y) && !double.IsInfinity(point.Y)) minimumY = Math.Min(minimumY, point.Y);
            }
        }
        return minimumY;
    }

    private static bool HasUsableVerticalPositioning(OfficeTextShapingResult run) {
        if (run.Direction != OfficeTextDirection.TopToBottom || run.Glyphs.Count == 0) return false;
        bool hasPenMovement = false;
        foreach (OfficeShapedGlyph glyph in run.Glyphs) {
            if (!glyph.AdvanceHeight.HasValue) return false;
            if (glyph.AdvanceHeight.Value != 0) hasPenMovement = true;
        }
        return hasPenMovement;
    }

    private double MeasureResolvedText(string text, IOfficeFontProgram font, double fontSize,
        OfficeTextFeatureSettings? featureSettings = null,
        OfficeTextDirection direction = OfficeTextDirection.Auto) {
        if (font.ProvidesComplexTextLayout) return font.Measure(text, fontSize);
        OfficeTextDirection resolvedDirection = ResolveTextDirection(text, direction);
        if (TryGetShapedTextRun(text, font, featureSettings, resolvedDirection, out OfficeTextShapingResult run)) {
            return font.MeasureShapedText(OfficeArabicTextShaper.ToLogicalText(text), run, fontSize);
        }
        OfficeManagedTextFallback fallback = GetManagedTextFallback(text, font, resolvedDirection);
        return font.Measure(fallback.Text, fontSize);
    }

    private List<List<OfficePoint>> GetResolvedTextContours(
        string text,
        IOfficeFontProgram font,
        double x,
        double y,
        double fontSize,
        OfficeTextFeatureSettings? featureSettings = null,
        OfficeTextDirection direction = OfficeTextDirection.Auto) {
        if (font.ProvidesComplexTextLayout) {
            return GetBoundedTextContours(font, text, x, y, fontSize);
        }
        OfficeTextDirection resolvedDirection = ResolveTextDirection(text, direction);
        if (TryGetShapedTextRun(text, font, featureSettings, resolvedDirection, out OfficeTextShapingResult run)) {
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
            GetManagedTextFallback(text, font, resolvedDirection).Text,
            x,
            y,
            fontSize);
    }

    private bool TryGetResolvedColorTextContours(
        string text,
        IOfficeFontProgram font,
        double x,
        double y,
        double fontSize,
        OfficeTextFeatureSettings? featureSettings,
        string? fontPalette,
        OfficeColor foreground,
        out List<OfficeColorGlyphContours> layers,
        OfficeTextDirection direction = OfficeTextDirection.Auto) {
        layers = new List<OfficeColorGlyphContours>();
        if (font is not OfficeTrueTypeFont trueType) return false;
        OfficeTextDirection resolvedDirection = ResolveTextDirection(text, direction);
        if (TryGetShapedTextRun(text, font, featureSettings, resolvedDirection, out OfficeTextShapingResult run)) {
            return trueType.TryGetShapedColorTextContours(
                OfficeArabicTextShaper.ToLogicalText(text),
                run,
                x,
                y,
                fontSize,
                fontPalette,
                foreground,
                MaximumTextOutlinePointsPerRun,
                _cancellationToken,
                out layers);
        }
        return trueType.TryGetColorTextContours(
            GetManagedTextFallback(text, font, resolvedDirection).Text,
            x,
            y,
            fontSize,
            fontPalette,
            foreground,
            MaximumTextOutlinePointsPerRun,
            _cancellationToken,
            out layers);
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

    private OfficeManagedTextFallback GetManagedTextFallback(string text, IOfficeFontProgram font,
        OfficeTextDirection direction = OfficeTextDirection.Auto) {
        _cancellationToken.ThrowIfCancellationRequested();
        OfficeTextDirection resolvedDirection = ResolveTextDirection(text, direction);
        var key = new ShapedTextKey(text, font, direction: resolvedDirection);
        Dictionary<ShapedTextKey, OfficeManagedTextFallback> cache =
            _managedTextCache ??= new Dictionary<ShapedTextKey, OfficeManagedTextFallback>();
        if (cache.TryGetValue(key, out OfficeManagedTextFallback cached)) return cached;

        OfficeManagedTextFallback fallback = OfficeManagedTextShaper.Resolve(
            text,
            font,
            resolvedDirection,
            _cancellationToken);
        if (fallback.Used || fallback.Incomplete) ReportTextShapingFallback(fallback.Incomplete);
        if (cache.Count >= MaxShapedTextCacheEntries) cache.Clear();
        cache[key] = fallback;
        return fallback;
    }

    private static OfficeTextDirection ResolveTextDirection(string text, OfficeTextDirection direction) =>
        direction == OfficeTextDirection.Auto
            ? OfficeTextElements.ResolveBaseDirection(OfficeArabicTextShaper.ToLogicalText(text))
            : direction;

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
        internal ShapedTextKey(string text, IOfficeFontProgram font, OfficeTextFeatureSettings? featureSettings = null, OfficeTextDirection direction = OfficeTextDirection.Auto) {
            Text = text;
            Font = font;
            FeatureSettings = featureSettings ?? OfficeTextFeatureSettings.Default;
            Direction = direction;
        }

        private string Text { get; }
        private IOfficeFontProgram Font { get; }
        private OfficeTextFeatureSettings FeatureSettings { get; }
        private OfficeTextDirection Direction { get; }

        public bool Equals(ShapedTextKey other) =>
            ReferenceEquals(Font, other.Font) &&
            Direction == other.Direction &&
            FeatureSettings.Equals(other.FeatureSettings) &&
            string.Equals(Text, other.Text, StringComparison.Ordinal);

        public override bool Equals(object? obj) =>
            obj is ShapedTextKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                return (StringComparer.Ordinal.GetHashCode(Text) * 397) ^
                       RuntimeHelpers.GetHashCode(Font) ^ FeatureSettings.GetHashCode() ^ (int)Direction;
            }
        }
    }
}
