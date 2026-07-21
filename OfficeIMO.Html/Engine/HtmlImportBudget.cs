namespace OfficeIMO.Html;

/// <summary>Tracks shared native-artifact import budgets for one conversion operation.</summary>
internal sealed class HtmlImportBudget {
    private readonly HtmlImportLimits _limits;
    private int _annotations;
    private long _chartPoints;
    private int _charts;
    private int _containers;
    private int _images;
    private long _imageBytes;
    private int _shapes;
    private int _tables;

    internal HtmlImportBudget(HtmlImportLimits limits) {
        _limits = (limits ?? throw new ArgumentNullException(nameof(limits))).Clone();
        _limits.Validate();
    }

    internal HtmlImportLimits Limits => _limits;

    internal bool TryReserveSemanticContainer(out string detail) =>
        TryIncrement(ref _containers, _limits.MaxSemanticContainers, nameof(HtmlImportLimits.MaxSemanticContainers), out detail);

    internal bool TryReserveTable(out string detail) =>
        TryIncrement(ref _tables, _limits.MaxTables, nameof(HtmlImportLimits.MaxTables), out detail);

    internal bool TryReserveShape(out string detail) =>
        TryIncrement(ref _shapes, _limits.MaxShapes, nameof(HtmlImportLimits.MaxShapes), out detail);

    internal bool TryReserveAnnotation(out string detail) =>
        TryIncrement(ref _annotations, _limits.MaxAnnotations, nameof(HtmlImportLimits.MaxAnnotations), out detail);

    internal bool TryReserveImage(HtmlImageDataUri dataUri, out string detail) {
        if (_images >= _limits.MaxImages) {
            detail = Detail(nameof(HtmlImportLimits.MaxImages), _images + 1L, _limits.MaxImages);
            return false;
        }

        long bytes;
        try {
            bytes = dataUri.EstimateDecodedByteCount();
        } catch (FormatException) {
            detail = "The embedded image payload was not valid base64 or percent-encoded data.";
            return false;
        }

        if (bytes <= 0L || bytes > _limits.MaxImageBytes) {
            detail = Detail(nameof(HtmlImportLimits.MaxImageBytes), bytes, _limits.MaxImageBytes);
            return false;
        }

        if (bytes > _limits.MaxTotalImageBytes - _imageBytes) {
            detail = Detail(nameof(HtmlImportLimits.MaxTotalImageBytes), _imageBytes + bytes, _limits.MaxTotalImageBytes);
            return false;
        }

        _images++;
        _imageBytes += bytes;
        detail = string.Empty;
        return true;
    }

    internal bool TryReserveChart(int series, int categories, out string detail) {
        if (_charts >= _limits.MaxCharts) {
            detail = Detail(nameof(HtmlImportLimits.MaxCharts), _charts + 1L, _limits.MaxCharts);
            return false;
        }

        if (series <= 0 || series > _limits.MaxChartSeries) {
            detail = Detail(nameof(HtmlImportLimits.MaxChartSeries), series, _limits.MaxChartSeries);
            return false;
        }

        if (categories <= 0 || categories > _limits.MaxChartCategories) {
            detail = Detail(nameof(HtmlImportLimits.MaxChartCategories), categories, _limits.MaxChartCategories);
            return false;
        }

        long points = (long)series * categories;
        if (points > _limits.MaxChartPoints - _chartPoints) {
            detail = Detail(nameof(HtmlImportLimits.MaxChartPoints), _chartPoints + points, _limits.MaxChartPoints);
            return false;
        }

        _charts++;
        _chartPoints += points;
        detail = string.Empty;
        return true;
    }

    internal bool IsMetadataWithinLimit(string? value, out string detail) {
        int length = value?.Length ?? 0;
        if (length <= _limits.MaxMetadataCharacters) {
            detail = string.Empty;
            return true;
        }

        detail = Detail(nameof(HtmlImportLimits.MaxMetadataCharacters), length, _limits.MaxMetadataCharacters);
        return false;
    }

    internal bool TryNormalizeGeometry(double value, double fallback, double minimum, out double normalized) {
        return TryNormalizeRange(value, fallback, minimum, _limits.MaxAbsoluteGeometry, out normalized);
    }

    internal bool TryNormalizeRange(double value, double fallback, double minimum, double maximum, out double normalized) {
        if (!double.IsNaN(value)
            && !double.IsInfinity(value)
            && value >= minimum
            && value <= maximum) {
            normalized = value;
            return true;
        }

        normalized = fallback;
        return false;
    }

    private static bool TryIncrement(ref int current, int limit, string source, out string detail) {
        if (current >= limit) {
            detail = Detail(source, current + 1L, limit);
            return false;
        }

        current++;
        detail = string.Empty;
        return true;
    }

    private static string Detail(string source, long actual, long limit) =>
        source + ": Actual=" + actual + "; Limit=" + limit;
}
