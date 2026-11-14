using System.Globalization;
using System.Net;
using System.Text;

namespace OfficeIMO.Diff;

/// <summary>Settings for <see cref="OfficeDiffHtmlReport"/>.</summary>
public sealed class OfficeDiffHtmlOptions {
    /// <summary>Page title. The default is "Comparison".</summary>
    public string Title { get; set; } = "Comparison";

    /// <summary>Heading of the old column. The default is "Old".</summary>
    public string OldTitle { get; set; } = "Old";

    /// <summary>Heading of the new column. The default is "New".</summary>
    public string NewTitle { get; set; } = "New";

    /// <summary>Unchanged lines kept around each change; longer unchanged stretches collapse to a marker row. A negative value shows every line. The default is 3.</summary>
    public int ContextLines { get; set; } = 3;

    /// <summary>Settings for the line and in-line word comparisons.</summary>
    public OfficeDiffOptions? DiffOptions { get; set; }

    /// <summary>Optional change list rendered under the side-by-side view, for example the result of a document redline.</summary>
    public OfficeChangeSet? Changes { get; set; }
}

/// <summary>Renders a standalone side-by-side HTML comparison of two texts.</summary>
/// <remarks>
/// Lines are aligned with a line diff. Replaced lines are paired and the words that differ inside them are highlighted.
/// The page has no scripts or external resources.
/// </remarks>
public static class OfficeDiffHtmlReport {
    /// <summary>Renders the comparison as a complete HTML document.</summary>
    public static string SideBySide(string? oldText, string? newText, OfficeDiffHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new OfficeDiffHtmlOptions();
        OfficeDiffResult<OfficeTextToken> lines = OfficeDiff.CompareLines(oldText, newText, options.DiffOptions, cancellationToken);
        var rows = BuildRows(lines, options, cancellationToken);

        var html = new StringBuilder();
        html.Append("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>")
            .Append(Encode(options.Title)).Append("</title>\n<style>\n").Append(Style).Append("</style>\n</head>\n<body>\n");
        html.Append("<h1>").Append(Encode(options.Title)).Append("</h1>\n");
        int added = lines.InsertedCount, removed = lines.DeletedCount;
        html.Append("<p class=\"summary\">").Append(removed.ToString(CultureInfo.InvariantCulture)).Append(" line(s) removed, ")
            .Append(added.ToString(CultureInfo.InvariantCulture)).Append(" line(s) added");
        if (lines.LimitReached) html.Append("; the comparison stopped at its work limit, so some unchanged lines may show as changed");
        html.Append(".</p>\n");

        html.Append("<table class=\"diff\">\n<colgroup><col class=\"num\"><col><col class=\"num\"><col></colgroup>\n<thead><tr><th colspan=\"2\">")
            .Append(Encode(options.OldTitle)).Append("</th><th colspan=\"2\">").Append(Encode(options.NewTitle)).Append("</th></tr></thead>\n<tbody>\n");
        if (lines.AreEqual) html.Append("<tr class=\"skip\"><td colspan=\"4\">The texts are identical.</td></tr>\n");
        foreach (string row in rows) html.Append(row);
        html.Append("</tbody>\n</table>\n");

        if (options.Changes != null) AppendChangeList(html, options.Changes);
        html.Append("</body>\n</html>\n");
        return html.ToString();
    }

    private static List<string> BuildRows(OfficeDiffResult<OfficeTextToken> lines, OfficeDiffHtmlOptions options, CancellationToken cancellationToken) {
        var rows = new List<string>();
        IReadOnlyList<OfficeDiffEdit> edits = lines.Edits;
        for (int editIndex = 0; editIndex < edits.Count; editIndex++) {
            OfficeDiffEdit edit = edits[editIndex];
            if (edit.Operation == OfficeDiffOperation.Equal) {
                bool first = editIndex == 0, last = editIndex == edits.Count - 1;
                int context = options.ContextLines;
                if (context < 0 || edit.OldLength <= (first || last ? context : context * 2)) {
                    for (int offset = 0; offset < edit.OldLength; offset++) rows.Add(EqualRow(lines, edit.OldStart + offset, edit.NewStart + offset));
                    continue;
                }

                int head = first ? 0 : context;
                int tail = last ? 0 : context;
                for (int offset = 0; offset < head; offset++) rows.Add(EqualRow(lines, edit.OldStart + offset, edit.NewStart + offset));
                rows.Add("<tr class=\"skip\"><td colspan=\"4\">" + (edit.OldLength - head - tail).ToString(CultureInfo.InvariantCulture) + " unchanged line(s)</td></tr>\n");
                for (int offset = edit.OldLength - tail; offset < edit.OldLength; offset++) rows.Add(EqualRow(lines, edit.OldStart + offset, edit.NewStart + offset));
                continue;
            }

            int oldStart = edit.OldStart, oldLength = 0, newStart = edit.NewStart, newLength = 0;
            while (editIndex < edits.Count && edits[editIndex].Operation != OfficeDiffOperation.Equal) {
                oldLength += edits[editIndex].OldLength;
                newLength += edits[editIndex].NewLength;
                editIndex++;
            }

            editIndex--;
            int paired = Math.Min(oldLength, newLength);
            for (int offset = 0; offset < Math.Max(oldLength, newLength); offset++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (offset < paired) {
                    rows.Add(ChangedRow(lines.Old[oldStart + offset], oldStart + offset, lines.New[newStart + offset], newStart + offset, options.DiffOptions, cancellationToken));
                } else if (offset < oldLength) {
                    rows.Add("<tr><td class=\"num\">" + (oldStart + offset + 1).ToString(CultureInfo.InvariantCulture) + "</td><td class=\"del\">" + Encode(Trim(lines.Old[oldStart + offset].Text)) + "</td><td class=\"num\"></td><td class=\"empty\"></td></tr>\n");
                } else {
                    rows.Add("<tr><td class=\"num\"></td><td class=\"empty\"></td><td class=\"num\">" + (newStart + offset + 1).ToString(CultureInfo.InvariantCulture) + "</td><td class=\"ins\">" + Encode(Trim(lines.New[newStart + offset].Text)) + "</td></tr>\n");
                }
            }
        }

        return rows;
    }

    private static string EqualRow(OfficeDiffResult<OfficeTextToken> lines, int oldIndex, int newIndex) =>
        "<tr><td class=\"num\">" + (oldIndex + 1).ToString(CultureInfo.InvariantCulture) + "</td><td>" + Encode(Trim(lines.Old[oldIndex].Text)) +
        "</td><td class=\"num\">" + (newIndex + 1).ToString(CultureInfo.InvariantCulture) + "</td><td>" + Encode(Trim(lines.New[newIndex].Text)) + "</td></tr>\n";

    private static string ChangedRow(OfficeTextToken oldLine, int oldIndex, OfficeTextToken newLine, int newIndex, OfficeDiffOptions? diffOptions, CancellationToken cancellationToken) {
        var wordOptions = diffOptions?.Clone() ?? new OfficeDiffOptions();
        wordOptions.SemanticCleanup = true;
        OfficeDiffResult<OfficeTextToken> words = OfficeDiff.CompareWords(Trim(oldLine.Text), Trim(newLine.Text), wordOptions, cancellationToken);
        var oldCell = new StringBuilder();
        var newCell = new StringBuilder();
        foreach (OfficeDiffEdit edit in words.Edits) {
            if (edit.Operation != OfficeDiffOperation.Insert) {
                string text = Concat(words.Old, edit.OldStart, edit.OldLength);
                if (edit.Operation == OfficeDiffOperation.Delete) oldCell.Append("<del>").Append(Encode(text)).Append("</del>");
                else oldCell.Append(Encode(text));
            }

            if (edit.Operation != OfficeDiffOperation.Delete) {
                string text = Concat(words.New, edit.NewStart, edit.NewLength);
                if (edit.Operation == OfficeDiffOperation.Insert) newCell.Append("<ins>").Append(Encode(text)).Append("</ins>");
                else newCell.Append(Encode(text));
            }
        }

        return "<tr><td class=\"num\">" + (oldIndex + 1).ToString(CultureInfo.InvariantCulture) + "</td><td class=\"chg\">" + oldCell +
               "</td><td class=\"num\">" + (newIndex + 1).ToString(CultureInfo.InvariantCulture) + "</td><td class=\"chg\">" + newCell + "</td></tr>\n";
    }

    private static void AppendChangeList(StringBuilder html, OfficeChangeSet changes) {
        html.Append("<h2>Changes</h2>\n");
        if (changes.IsEmpty) {
            html.Append("<p class=\"summary\">No changes.</p>\n");
            return;
        }

        html.Append("<table class=\"changes\">\n<thead><tr><th>#</th><th>Kind</th><th>Old page</th><th>Old text</th><th>New page</th><th>New text</th></tr></thead>\n<tbody>\n");
        foreach (OfficeChange change in changes.Changes) {
            html.Append("<tr><td>").Append(change.Id.ToString(CultureInfo.InvariantCulture)).Append("</td><td>").Append(Encode(change.Kind.ToString()))
                .Append("</td><td>").Append(Page(change.Old)).Append("</td><td class=\"del\">").Append(Encode(change.OldText))
                .Append("</td><td>").Append(Page(change.New)).Append("</td><td class=\"ins\">").Append(Encode(change.NewText)).Append("</td></tr>\n");
        }

        html.Append("</tbody>\n</table>\n");
    }

    private static string Page(OfficeChangeAnchor? anchor) => anchor?.PageNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;

    private static string Concat(IReadOnlyList<OfficeTextToken> tokens, int start, int length) {
        var builder = new StringBuilder();
        for (int index = start; index < start + length; index++) builder.Append(tokens[index].Text);
        return builder.ToString();
    }

    private static string Trim(string line) => line.TrimEnd('\r', '\n');

    private static string Encode(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);

    private const string Style =
        ":root{color-scheme:light dark;--ins:#d7f5dd;--del:#fbd9d9;--chg:#fff8e1;--num:#8a8a8a;--line:#e3e3e3;--skip:#f2f2f2}\n" +
        "@media (prefers-color-scheme:dark){:root{--ins:#1f4a2a;--del:#5a2323;--chg:#3d3720;--num:#9a9a9a;--line:#3a3a3a;--skip:#2a2a2a}}\n" +
        "body{font-family:system-ui,-apple-system,Segoe UI,sans-serif;margin:16px;line-height:1.4}\n" +
        "table{border-collapse:collapse;width:100%;table-layout:fixed}\n" +
        "th,td{border-bottom:1px solid var(--line);padding:2px 6px;vertical-align:top;text-align:left}\n" +
        "table.diff td{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;font-size:13px;white-space:pre-wrap;word-break:break-word}\n" +
        "col.num{width:3.5em}td.num{color:var(--num);text-align:right;user-select:none}\n" +
        "td.ins,ins{background:var(--ins);text-decoration:none}td.del,del{background:var(--del)}td.chg{background:var(--chg)}\n" +
        "tr.skip td{background:var(--skip);color:var(--num);text-align:center;font-style:italic}\n" +
        "table.changes{margin-top:8px}table.changes td{white-space:pre-wrap;word-break:break-word}\n" +
        "p.summary{color:var(--num)}\n";
}
