using System.Globalization;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    private sealed class TablePart {
        internal string Title = "";
        internal string[] Headers = Array.Empty<string>();
        internal string[][] Rows = Array.Empty<string[]>();
    }

    private static bool IsUsage(ProjectView view) => view.Kind == ProjectViewKind.TaskUsage || view.Kind == ProjectViewKind.ResourceUsage || view.Kind == ProjectViewKind.ResourceHistogram;

    private static int[] ColumnWeights(string[] headers) => headers.Select(h => h.Contains("Name") ? 4
        : h == "UID" || h == "Critical" || h == "Summary" ? 1 : 3).ToArray();

    private static IEnumerable<TablePart> TableParts(ProjectView view, CancellationToken token) {
        const int maxColumns = 6;
        for (int c = 0; c < view.Columns.Count; c += maxColumns - 1) {
            var columns = view.Columns.Skip(c).Take(maxColumns - 1).ToList();
            // Keep horizontal continuations independently identifiable even when the caller omitted UID.
            if (view.Columns.Count > maxColumns - 1 && !columns.Contains(ProjectViewColumn.Uid)) columns.Insert(0, ProjectViewColumn.Uid);
            foreach (var part in SplitRows(view.Kind.ToString(), columns.Select(ProjectView.ColumnTitle).ToArray(),
                view.Rows.Select(row => columns.Select(row.GetText).ToArray()), token)) yield return part;
        }
        if (IsUsage(view)) for (int b = 0; b < view.Buckets.Count; b += maxColumns - 1) {
            int offset = b, count = Math.Min(maxColumns - 1, view.Buckets.Count - b);
            foreach (var part in SplitRows("Work hours · " + view.Buckets[b].Start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                new[] { "UID / Name" }.Concat(view.Buckets.Skip(b).Take(count).Select(bucket => bucket.Start.ToString("dd MMM yy", CultureInfo.InvariantCulture))).ToArray(),
                view.Rows.Select(row => new[] { row.Uid.ToString(CultureInfo.InvariantCulture) + " / " + row.Name }
                    .Concat(row.BucketWorkHours.Skip(offset).Take(count).Select(value => value.ToString("0.##", CultureInfo.InvariantCulture))).ToArray()), token)) yield return part;
        }
        foreach (var part in SplitRows("Status and groups", new[] { "UID", "Name", "Status", "Group" },
            view.Rows.Select(row => new[] { row.Uid.ToString(CultureInfo.InvariantCulture), row.Name, Status(view, row), row.Group }), token)) yield return part;
        if (view.Rows.Any(r => r.BaselineStart.HasValue || r.BaselineFinish.HasValue)) {
            foreach (var part in SplitRows("Baseline dates", new[] { "UID", "Name", "Baseline start", "Baseline finish" },
                view.Rows.Select(row => new[] { row.Uid.ToString(CultureInfo.InvariantCulture), row.Name, DateText(row.BaselineStart), DateText(row.BaselineFinish) }), token)) yield return part;
        }
        if (view.Links.Count > 0) {
            foreach (var part in SplitRows("Dependencies", new[] { "Predecessor UID", "Successor UID", "Type", "Lag" },
                view.Links.Select(link => new[] { link.PredecessorUid.ToString(CultureInfo.InvariantCulture), link.SuccessorUid.ToString(CultureInfo.InvariantCulture), link.Type.ToString(), link.LagText }), token)) yield return part;
        }
    }

    private static IEnumerable<TablePart> SplitRows(string title, string[] headers, IEnumerable<string[]> values, CancellationToken token) {
        var rows = new List<string[]>(14); int offset = 0;
        foreach (var row in values) {
            token.ThrowIfCancellationRequested(); rows.Add(row);
            if (rows.Count == 14) {
                yield return new TablePart { Title = title + " · rows " + (offset + 1) + "–" + (offset + rows.Count), Headers = headers, Rows = rows.ToArray() };
                offset += rows.Count; rows.Clear();
            }
        }
        if (rows.Count > 0 || offset == 0) yield return new TablePart {
            Title = title + (rows.Count == 0 ? " · no matching rows" : " · rows " + (offset + 1) + "–" + (offset + rows.Count)),
            Headers = headers, Rows = rows.ToArray()
        };
    }

    private static string DateText(DateTime? value) => value?.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) ?? "";
    private static string Status(ProjectView view, ProjectViewRow row) =>
        (view.Kind == ProjectViewKind.ResourceUsage || view.Kind == ProjectViewKind.ResourceHistogram ? "Resource" : row.IsSummary ? "Summary" : "Task")
        + (row.IsCritical ? "; critical" : "");
}
