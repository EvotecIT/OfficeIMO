namespace OfficeIMO.Project;

internal static class ProjectViewBuilder {
    internal static bool Matches(string? name, string? filter) => string.IsNullOrEmpty(filter) || (name ?? "").IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;

    internal static ProjectViewOptions CopyOptions(ProjectViewOptions input) {
        if (!Enum.IsDefined(typeof(ProjectViewKind), input.Kind) || !Enum.IsDefined(typeof(ProjectViewTimescale), input.Timescale)
            || !Enum.IsDefined(typeof(ProjectViewGrouping), input.Grouping)
            || !Enum.IsDefined(typeof(ProjectNetworkLayout), input.NetworkLayout)) throw new ArgumentException("Unknown report layout option.");
        if (input.StatusDate.HasValue && input.StatusDate.Value.Kind != DateTimeKind.Unspecified)
            throw new ArgumentException("StatusDate must use DateTimeKind.Unspecified.", nameof(input));
        if (input.MaxRows < 1 || input.MaxBuckets < 1 || input.MaxCells < 1 || input.MaxPages < 1 || input.MaxIntervalVisits < 1) throw new ArgumentOutOfRangeException(nameof(input), "Report limits must be positive.");
        if (!Finite(input.PageWidth) || !Finite(input.PageHeight) || !Finite(input.Margin) || input.Margin < 12
            || input.PageWidth - 2 * input.Margin < 300 || input.PageHeight - 2 * input.Margin < 180
            || input.PageWidth > 14400 || input.PageHeight > 14400) throw new ArgumentOutOfRangeException(nameof(input), "Invalid page dimensions or margins.");
        if (input.BaselineNumber < 0 || input.BaselineNumber > 10) throw new ArgumentOutOfRangeException(nameof(input.BaselineNumber));
        if (input.Start.HasValue && input.Start.Value.Kind != DateTimeKind.Unspecified || input.Finish.HasValue && input.Finish.Value.Kind != DateTimeKind.Unspecified)
            throw new ArgumentException("Visible report dates must use DateTimeKind.Unspecified.", nameof(input));
        if (input.Start.HasValue && input.Finish.HasValue && input.Start >= input.Finish) throw new ArgumentException("Finish must follow Start.");
        return new ProjectViewOptions { Kind = input.Kind, Timescale = input.Timescale, IncludeSummaries = input.IncludeSummaries,
            NetworkLayout = input.NetworkLayout, StatusDate = input.StatusDate, ShowProgress = input.ShowProgress,
            CriticalOnly = input.CriticalOnly, NameContains = input.NameContains, Grouping = input.Grouping, BaselineNumber = input.BaselineNumber,
            Start = input.Start, Finish = input.Finish, PageWidth = input.PageWidth, PageHeight = input.PageHeight,
            FitPageHeightToContent = input.FitPageHeightToContent,
            Margin = input.Margin, ShowLegend = input.ShowLegend, MaxRows = input.MaxRows, MaxBuckets = input.MaxBuckets,
            MaxCells = input.MaxCells, MaxPages = input.MaxPages, MaxIntervalVisits = input.MaxIntervalVisits };
    }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    internal static HashSet<int>? SelectIds(IReadOnlyCollection<int>? input, IEnumerable<int> available, int limit, string name) {
        if (input == null) return null;
        var result = new HashSet<int>(); var valid = new HashSet<int>(available); int count = 0;
        foreach (int uid in input) {
            if (++count > limit) throw new ArgumentException("Selection exceeds MaxRows.", name);
            if (!valid.Contains(uid)) throw new ArgumentException("Selection contains an unknown UID: " + uid, name);
            result.Add(uid);
        }
        return result;
    }

    internal static void CheckCells(int rows, int buckets, ProjectViewOptions options) {
        if (rows > options.MaxRows || (long)rows * buckets > options.MaxCells) throw new InvalidOperationException("Report exceeds its row or cell limit.");
    }

    internal static ProjectViewBucket[] Buckets(ProjectViewOptions options, (DateTime Start, DateTime Finish)[] ranges) {
        if (options.Kind == ProjectViewKind.Table || options.Kind == ProjectViewKind.Network) return Array.Empty<ProjectViewBucket>();
        if (ranges.Length == 0 && (!options.Start.HasValue || !options.Finish.HasValue)) return Array.Empty<ProjectViewBucket>();
        var start = options.Start ?? ranges.Min(t => t.Start).Date;
        if (!options.Start.HasValue && options.Timescale == ProjectViewTimescale.Week)
            start = start.AddDays(-((int)start.DayOfWeek + 6) % 7);
        if (!options.Start.HasValue && options.Timescale == ProjectViewTimescale.Month)
            start = new DateTime(start.Year, start.Month, 1);
        var finish = options.Finish ?? ranges.Max(t => t.Finish);
        // A zero-duration event at an inferred end needs space inside the exclusive range.
        // Leave explicit caller clipping unchanged and retain a visible marker at midnight too.
        if (!options.Finish.HasValue && ranges.Any(t => t.Start == finish && t.Finish == finish)) {
            if (finish.Date == DateTime.MaxValue.Date) throw new ArgumentException("A terminal milestone at the maximum date needs an explicit report range.");
            finish = finish.Date.AddDays(1);
        }
        if (finish == start && !options.Finish.HasValue) finish = start.AddDays(1);
        if (finish <= start) throw new ArgumentException("The visible report range is empty or reversed.");
        var result = new List<ProjectViewBucket>();
        while (start < finish) {
            if (result.Count >= options.MaxBuckets) throw new InvalidOperationException("Report exceeds MaxBuckets.");
            DateTime next;
            try {
                next = options.Timescale == ProjectViewTimescale.Month ? new DateTime(start.Year, start.Month, 1).AddMonths(1)
                    : start.Date.AddDays(options.Timescale == ProjectViewTimescale.Week ? 7 - ((int)start.DayOfWeek + 6) % 7 : 1);
            } catch (ArgumentOutOfRangeException) { next = DateTime.MaxValue; }
            if (next > finish) next = finish;
            result.Add(new ProjectViewBucket(start, next)); start = next;
        }
        return result.ToArray();
    }

    internal static decimal[] WorkBuckets(IEnumerable<ProjectAssignmentSchedule> assignments, ProjectViewBucket[] buckets, CancellationToken token, ref long visits, int maxVisits) {
        var result = new decimal[buckets.Length];
        if (buckets.Length == 0) return result;
        foreach (var assignment in assignments) foreach (var interval in assignment.Intervals) {
            token.ThrowIfCancellationRequested();
            if (++visits > maxVisits) throw new InvalidOperationException("Report exceeds MaxIntervalVisits.");
            if (interval.Finish <= interval.Start) continue;
            // Locate the first intersecting bucket without scanning unrelated dates.
            int low = 0, high = buckets.Length;
            while (low < high) { int middle = low + (high - low) / 2; if (buckets[middle].Finish <= interval.Start) low = middle + 1; else high = middle; }
            for (int i = low; i < buckets.Length && buckets[i].Start < interval.Finish; i++) {
                if (++visits > maxVisits) throw new InvalidOperationException("Report exceeds MaxIntervalVisits.");
                var from = interval.Start > buckets[i].Start ? interval.Start : buckets[i].Start;
                var to = interval.Finish < buckets[i].Finish ? interval.Finish : buckets[i].Finish;
                if (to > from) result[i] += interval.Work.Minutes / 60m * (to.Ticks - from.Ticks) / (interval.Finish.Ticks - interval.Start.Ticks);
            }
        }
        return result;
    }
}
