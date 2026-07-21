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

    internal bool TryReserveTableWithShape(out string detail) {
        if (!CanIncrement(_tables, _limits.MaxTables, nameof(HtmlImportLimits.MaxTables), out detail)
            || !CanIncrement(_shapes, _limits.MaxShapes, nameof(HtmlImportLimits.MaxShapes), out detail)) {
            return false;
        }

        _tables++;
        _shapes++;
        detail = string.Empty;
        return true;
    }

    internal bool TryReserveShape(out string detail) =>
        TryIncrement(ref _shapes, _limits.MaxShapes, nameof(HtmlImportLimits.MaxShapes), out detail);

    internal bool TryReserveAnnotation(out string detail) =>
        TryIncrement(ref _annotations, _limits.MaxAnnotations, nameof(HtmlImportLimits.MaxAnnotations), out detail);

    internal bool IsImageWithinLimit(HtmlImageDataUri dataUri, out string detail) =>
        TryGetImageByteCount(dataUri, out _, out detail);

    internal bool TryReserveImageWithShape(HtmlImageDataUri dataUri, out string detail) {
        if (!TryGetImageByteCount(dataUri, out long bytes, out detail)
            || !CanIncrement(_shapes, _limits.MaxShapes, nameof(HtmlImportLimits.MaxShapes), out detail)) {
            return false;
        }

        _images++;
        _imageBytes += bytes;
        _shapes++;
        detail = string.Empty;
        return true;
    }

    internal bool TryReserveChartWithShape(
        int series,
        int categories,
        out HtmlImportBudgetReservation reservation,
        out string detail) {
        reservation = null!;
        if (!CanReserveChart(series, categories, out long points, out detail)
            || !CanIncrement(_shapes, _limits.MaxShapes, nameof(HtmlImportLimits.MaxShapes), out detail)) {
            return false;
        }

        _charts++;
        _chartPoints += points;
        _shapes++;
        reservation = new HtmlImportBudgetReservation(() => ReleaseChartWithShape(points));
        detail = string.Empty;
        return true;
    }

    private void ReleaseChartWithShape(long points) {
        _charts--;
        _chartPoints -= points;
        _shapes--;
    }

    private bool CanReserveChart(int series, int categories, out long points, out string detail) {
        if (_charts >= _limits.MaxCharts) {
            points = 0L;
            detail = Detail(nameof(HtmlImportLimits.MaxCharts), _charts + 1L, _limits.MaxCharts);
            return false;
        }

        if (series <= 0 || series > _limits.MaxChartSeries) {
            points = 0L;
            detail = Detail(nameof(HtmlImportLimits.MaxChartSeries), series, _limits.MaxChartSeries);
            return false;
        }

        if (categories <= 0 || categories > _limits.MaxChartCategories) {
            points = 0L;
            detail = Detail(nameof(HtmlImportLimits.MaxChartCategories), categories, _limits.MaxChartCategories);
            return false;
        }

        points = (long)series * categories;
        if (points > _limits.MaxChartPoints - _chartPoints) {
            detail = Detail(nameof(HtmlImportLimits.MaxChartPoints), _chartPoints + points, _limits.MaxChartPoints);
            return false;
        }

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
        if (!CanIncrement(current, limit, source, out detail)) return false;

        current++;
        detail = string.Empty;
        return true;
    }

    private bool TryGetImageByteCount(HtmlImageDataUri dataUri, out long bytes, out string detail) {
        if (_images >= _limits.MaxImages) {
            bytes = 0L;
            detail = Detail(nameof(HtmlImportLimits.MaxImages), _images + 1L, _limits.MaxImages);
            return false;
        }

        try {
            bytes = dataUri.EstimateDecodedByteCount();
        } catch (FormatException) {
            bytes = 0L;
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

        detail = string.Empty;
        return true;
    }

    private static bool CanIncrement(int current, int limit, string source, out string detail) {
        if (current < limit) {
            detail = string.Empty;
            return true;
        }

        detail = Detail(source, current + 1L, limit);
        return false;
    }

    private static string Detail(string source, long actual, long limit) =>
        source + ": Actual=" + actual + "; Limit=" + limit;
}

/// <summary>Rolls back a provisional shared import-budget reservation unless it is committed.</summary>
internal sealed class HtmlImportBudgetReservation : IDisposable {
    private Action? _rollback;

    internal HtmlImportBudgetReservation(Action rollback) {
        _rollback = rollback ?? throw new ArgumentNullException(nameof(rollback));
    }

    internal void Commit() => _rollback = null;

    public void Dispose() {
        Action? rollback = _rollback;
        _rollback = null;
        rollback?.Invoke();
    }
}
