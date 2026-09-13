namespace OfficeIMO.Project;

/// <summary>One calculation's sorted, non-overlapping resource capacity and cost tables.</summary>
internal sealed class ProjectResourceTimeline {
    private readonly ProjectResource _resource;
    private readonly ProjectResourceAvailability[] _availability;
    private readonly Dictionary<ProjectCostRateTable, ProjectResourceRate[]> _rates;
    internal ProjectResourceTimeline(ProjectResource resource) {
        _resource = resource;
        _availability = resource.AvailabilityPeriods.OrderBy(p => p.From ?? DateTime.MinValue).ToArray();
        _rates = resource.Rates.GroupBy(r => r.Table ?? ProjectCostRateTable.A).ToDictionary(g => g.Key, g => g.OrderBy(r => r.From).ToArray());
    }
    internal decimal Capacity(DateTime at) {
        if (_availability.Length == 0) return _resource.MaxUnits?.Value ?? 1m;
        int index = Find(_availability.Length, i => _availability[i].From ?? DateTime.MinValue, at);
        if (index < 0) return 0m;
        var period = _availability[index];
        return !period.Through.HasValue || at < AvailabilityEnd(period.Through.Value) ? period.Units?.Value ?? _resource.MaxUnits?.Value ?? 1m : 0m;
    }
    internal static DateTime AvailabilityEnd(DateTime through) => through > DateTime.MaxValue.AddMinutes(-1) ? DateTime.MaxValue : through.AddMinutes(1);
    internal IEnumerable<DateTime> CapacityBoundaries(DateTime start, DateTime finish) {
        foreach (var period in _availability) {
            if (period.From > start && period.From < finish) yield return period.From.Value;
            if (period.Through.HasValue) { var end = AvailabilityEnd(period.Through.Value); if (end > start && end < finish) yield return end; }
        }
    }
    internal (decimal? Standard, decimal? Overtime, decimal? PerUse) Rate(DateTime at, ProjectCostRateTable table) {
        if (!_rates.TryGetValue(table, out var rates))
            return table == ProjectCostRateTable.A ? (_resource.StandardRate, _resource.OvertimeRate, _resource.CostPerUse) : (null, null, null);
        int index = Find(rates.Length, i => rates[i].From ?? DateTime.MinValue, at);
        if (index < 0 || !rates[index].To.HasValue || at >= rates[index].To) return (null, null, null);
        var rate = rates[index]; return (rate.StandardRate, rate.OvertimeRate, rate.CostPerUse);
    }
    internal IEnumerable<DateTime> RateBoundaries(DateTime start, DateTime finish, ProjectCostRateTable table) {
        if (!_rates.TryGetValue(table, out var rates)) yield break;
        foreach (var rate in rates) {
            if (rate.From > start && rate.From < finish) yield return rate.From.Value;
            if (rate.To > start && rate.To < finish) yield return rate.To.Value;
        }
    }
    private static int Find(int count, Func<int, DateTime> start, DateTime at) {
        int low = 0, high = count - 1, result = -1;
        while (low <= high) { int middle = low + (high - low) / 2; if (start(middle) <= at) { result = middle; low = middle + 1; } else high = middle - 1; }
        return result;
    }
}
