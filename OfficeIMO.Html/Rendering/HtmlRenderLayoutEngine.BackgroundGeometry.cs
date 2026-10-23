using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool TryResolveBackgroundRepeat(string value, out BackgroundRepeatMode repeatX, out BackgroundRepeatMode repeatY) {
        repeatX = BackgroundRepeatMode.Repeat;
        repeatY = BackgroundRepeatMode.Repeat;
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(value ?? string.Empty)
            .Select(token => token.Trim().ToLowerInvariant())
            .Where(token => token.Length > 0)
            .ToList()
            .AsReadOnly();
        if (values.Count == 0) return true;
        if (values.Count > 2) return false;
        if (values.Count == 1 && values[0] == "repeat-x") {
            repeatX = BackgroundRepeatMode.Repeat;
            repeatY = BackgroundRepeatMode.NoRepeat;
            return true;
        }

        if (values.Count == 1 && values[0] == "repeat-y") {
            repeatX = BackgroundRepeatMode.NoRepeat;
            repeatY = BackgroundRepeatMode.Repeat;
            return true;
        }

        if (!TryParseBackgroundRepeatMode(values[0], out repeatX)) return false;
        if (values.Count == 1) {
            repeatY = repeatX;
            return true;
        }

        return TryParseBackgroundRepeatMode(values[1], out repeatY);
    }

    private static bool TryParseBackgroundRepeatMode(string value, out BackgroundRepeatMode mode) {
        switch (value) {
            case "repeat":
                mode = BackgroundRepeatMode.Repeat;
                return true;
            case "no-repeat":
                mode = BackgroundRepeatMode.NoRepeat;
                return true;
            case "space":
                mode = BackgroundRepeatMode.Space;
                return true;
            case "round":
                mode = BackgroundRepeatMode.Round;
                return true;
            default:
                mode = BackgroundRepeatMode.NoRepeat;
                return false;
        }
    }

    private static BackgroundRepeatAxis ResolveBackgroundRepeatAxis(
        BackgroundRepeatMode mode,
        double areaStart,
        double areaLength,
        double positionedTileStart,
        double tileLength) {
        if (mode == BackgroundRepeatMode.NoRepeat) return new BackgroundRepeatAxis(positionedTileStart, tileLength, repeat: false);
        if (mode == BackgroundRepeatMode.Repeat || mode == BackgroundRepeatMode.Round)
            return new BackgroundRepeatAxis(positionedTileStart, tileLength, repeat: true);

        double availableCount = Math.Floor((areaLength + 0.0000001D) / tileLength);
        long count = availableCount >= long.MaxValue ? long.MaxValue : Math.Max(0L, (long)availableCount);
        if (count < 2L) return new BackgroundRepeatAxis(positionedTileStart, tileLength, repeat: false);
        double step = (areaLength - tileLength) / (count - 1L);
        return new BackgroundRepeatAxis(areaStart, Math.Max(tileLength, step), repeat: true);
    }

    private static BackgroundImageSize ApplyRoundBackgroundSize(
        BackgroundImageSize size,
        string sizeValue,
        double areaWidth,
        double areaHeight,
        BackgroundRepeatMode repeatX,
        BackgroundRepeatMode repeatY,
        bool hasIntrinsicRatio = true) {
        double width = size.Width;
        double height = size.Height;
        double originalWidth = width;
        double originalHeight = height;
        ResolveBackgroundSizeAutoAxes(sizeValue, out bool widthAuto, out bool heightAuto);
        if (repeatX == BackgroundRepeatMode.Round) {
            width = ResolveRoundedTileLength(areaWidth, width);
            if (hasIntrinsicRatio && repeatY != BackgroundRepeatMode.Round && heightAuto) height *= width / Math.Max(0.01D, originalWidth);
        }

        if (repeatY == BackgroundRepeatMode.Round) {
            height = ResolveRoundedTileLength(areaHeight, height);
            if (hasIntrinsicRatio && repeatX != BackgroundRepeatMode.Round && widthAuto) width *= height / Math.Max(0.01D, originalHeight);
        }

        return new BackgroundImageSize(Math.Max(0D, width), Math.Max(0D, height));
    }

    private static double ResolveRoundedTileLength(double areaLength, double tileLength) {
        double ratio = areaLength / tileLength;
        long count = ratio >= long.MaxValue ? long.MaxValue
            : Math.Max(1L, (long)Math.Round(ratio, MidpointRounding.AwayFromZero));
        return areaLength / count;
    }

    private static void ResolveBackgroundSizeAutoAxes(string value, out bool widthAuto, out bool heightAuto) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value ?? string.Empty);
        if (parts.Count == 0 || (parts.Count == 1 && string.Equals(parts[0], "auto", StringComparison.OrdinalIgnoreCase))) {
            widthAuto = true;
            heightAuto = true;
            return;
        }

        if (parts.Count == 1
            && (string.Equals(parts[0], "cover", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parts[0], "contain", StringComparison.OrdinalIgnoreCase))) {
            widthAuto = false;
            heightAuto = false;
            return;
        }

        widthAuto = string.Equals(parts[0], "auto", StringComparison.OrdinalIgnoreCase);
        heightAuto = parts.Count == 1 || string.Equals(parts[1], "auto", StringComparison.OrdinalIgnoreCase);
    }

    private BackgroundImageSize ResolveBackgroundImageSize(
        string value,
        double areaWidth,
        double areaHeight,
        double intrinsicWidth,
        double intrinsicHeight,
        HtmlRenderBoxStyle style,
        out bool usedFallback,
        bool hasIntrinsicRatio = true) {
        usedFallback = false;
        string normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "auto") {
            return new BackgroundImageSize(intrinsicWidth, intrinsicHeight);
        }

        if (normalized == "cover" || normalized == "contain") {
            if (!hasIntrinsicRatio) return new BackgroundImageSize(areaWidth, areaHeight);
            double scaleX = areaWidth / Math.Max(0.01D, intrinsicWidth);
            double scaleY = areaHeight / Math.Max(0.01D, intrinsicHeight);
            double scale = normalized == "cover" ? Math.Max(scaleX, scaleY) : Math.Min(scaleX, scaleY);
            return new BackgroundImageSize(intrinsicWidth * scale, intrinsicHeight * scale);
        }

        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(normalized).ToList().AsReadOnly();
        if (parts.Count == 0 || parts.Count > 2) {
            usedFallback = true;
            return Contain(areaWidth, areaHeight, intrinsicWidth, intrinsicHeight);
        }

        bool widthAuto = parts[0] == "auto";
        bool heightAuto = parts.Count == 1 || parts[1] == "auto";
        double resolvedWidth = intrinsicWidth;
        double resolvedHeight = intrinsicHeight;
        if (!widthAuto && !TryResolveLength(parts[0], areaWidth, style, out resolvedWidth)) {
            usedFallback = true;
            return Contain(areaWidth, areaHeight, intrinsicWidth, intrinsicHeight);
        }

        if (!heightAuto && !TryResolveLength(parts[1], areaHeight, style, out resolvedHeight)) {
            usedFallback = true;
            return Contain(areaWidth, areaHeight, intrinsicWidth, intrinsicHeight);
        }

        double ratio = intrinsicWidth / Math.Max(0.01D, intrinsicHeight);
        if (hasIntrinsicRatio && widthAuto && !heightAuto) resolvedWidth = resolvedHeight * ratio;
        if (hasIntrinsicRatio && !widthAuto && heightAuto) resolvedHeight = resolvedWidth / Math.Max(0.01D, ratio);
        return new BackgroundImageSize(Math.Max(0D, resolvedWidth), Math.Max(0D, resolvedHeight));
    }

    private static BackgroundImageSize Contain(double areaWidth, double areaHeight, double intrinsicWidth, double intrinsicHeight) {
        double scale = Math.Min(areaWidth / Math.Max(0.01D, intrinsicWidth), areaHeight / Math.Max(0.01D, intrinsicHeight));
        return new BackgroundImageSize(intrinsicWidth * scale, intrinsicHeight * scale);
    }

    /// <summary>Resolves the shared CSS position grammar against the concrete image's remaining space.</summary>
    private (double X, double Y) ResolveBackgroundPosition(
        string value, double areaWidth, double areaHeight, BackgroundImageSize size, HtmlRenderBoxStyle style) {
        HtmlCssReplacedElementParser.TryResolveObjectPosition(
            value.Trim().ToLowerInvariant(), areaWidth, areaHeight, size.Width, size.Height,
            style.Font.Size, _styleResolver.RootFontSize, ActiveSurfaceWidth,
            _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
            style.ContainerUnitWidth ?? double.NaN, style.ContainerUnitHeight ?? double.NaN,
            out double x, out double y, style.CharacterAdvance);
        return (x, y);
    }

    private void ResolveBackgroundRepeatModes(
        string value, string source, out BackgroundRepeatMode horizontal, out BackgroundRepeatMode vertical) {
        if (TryResolveBackgroundRepeat(value, out horizontal, out vertical)) return;
        AddUnsupported(HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported,
            "A CSS background-repeat value used a single-image fallback.", source, value);
        horizontal = BackgroundRepeatMode.NoRepeat;
        vertical = BackgroundRepeatMode.NoRepeat;
    }

    /// <summary>Positions and distributes tiles in the origin while retaining the independent painting clip.</summary>
    private OfficeImagePatternLayout ResolveBackgroundPattern(
        HtmlRenderBackgroundLayer layer, HtmlRenderBoxStyle style,
        BackgroundPaintBox origin, BackgroundPaintBox clip, BackgroundImageSize size,
        BackgroundRepeatMode repeatX, BackgroundRepeatMode repeatY) {
        (double x, double y) = ResolveBackgroundPosition(layer.Position, origin.Width, origin.Height, size, style);
        BackgroundRepeatAxis horizontal = ResolveBackgroundRepeatAxis(repeatX, origin.X, origin.Width, origin.X + x, size.Width);
        BackgroundRepeatAxis vertical = ResolveBackgroundRepeatAxis(repeatY, origin.Y, origin.Height, origin.Y + y, size.Height);
        return new OfficeImagePatternLayout(
            new OfficeImagePlacement(clip.X, clip.Y, clip.Width, clip.Height),
            new OfficeImagePlacement(horizontal.Origin, vertical.Origin, size.Width, size.Height),
            horizontal.Repeat, vertical.Repeat, horizontal.Step, vertical.Step);
    }

    private bool TryReserveBackgroundTiles(long count, string source) {
        CheckCancellation();
        if (_operationBudget.TryReserveBackgroundImageTiles(count, _options.MaxBackgroundImageTiles)) return true;
        if (count > 0L) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded,
                "Repeated CSS background images exceeded the configured operation-wide tile limit and used one clipped origin tile.",
                HtmlDiagnosticSeverity.Error, source,
                "tiles=" + count.ToString(CultureInfo.InvariantCulture) + ";limit=" + _options.MaxBackgroundImageTiles.ToString(CultureInfo.InvariantCulture));
        }
        return false;
    }

    private readonly struct BackgroundPaintBox {
        internal BackgroundPaintBox(double x, double y, double width, double height, HtmlResolvedBorderRadii radii) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Radii = radii;
        }

        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal HtmlResolvedBorderRadii Radii { get; }
    }

    private readonly struct BackgroundImageSize {
        internal BackgroundImageSize(double width, double height) {
            Width = width;
            Height = height;
        }

        internal double Width { get; }
        internal double Height { get; }
    }

    private readonly struct BackgroundRepeatAxis {
        internal BackgroundRepeatAxis(double origin, double step, bool repeat) {
            Origin = origin;
            Step = step;
            Repeat = repeat;
        }

        internal double Origin { get; }
        internal double Step { get; }
        internal bool Repeat { get; }
    }

    private enum BackgroundRepeatMode {
        NoRepeat,
        Repeat,
        Space,
        Round
    }
}
