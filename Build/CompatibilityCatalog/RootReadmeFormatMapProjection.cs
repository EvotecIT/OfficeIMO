using System.Globalization;
using System.Text;

/// <summary>
/// Projects the website format map into the root README: one row per format family with the formats and the number of
/// conversion routes that start there, plus how many routes each surface runs. The block sits between standalone markers
/// and is replaced in place; the rest of the README stays hand-written.
/// </summary>
internal static class RootReadmeFormatMapProjection {
    internal const string StartMarker = "<!-- officeimo-format-map:start -->";
    internal const string EndMarker = "<!-- officeimo-format-map:end -->";

    /// <summary>
    /// README rows. The website keeps less common formats in its "other" family to keep the circle readable; the README
    /// has room to place them with the family a reader looks for. A format not named here keeps its website family, and
    /// anything still unplaced lands in the last row, so a new format is never dropped.
    /// </summary>
    private static readonly (string Title, string[] WebsiteFamilies, string[] Formats)[] Rows = [
        ("Documents", ["documents"], ["DOCM", "DOTX", "DOTM", "Word XML", "Word"]),
        ("Spreadsheets", ["spreadsheets"], ["XLSB", "XLSM", "XLS", "XLTX", "XLTM", "XLAM", "XML Spreadsheet 2003", "TSV", "Excel"]),
        ("Presentations", ["presentations"], ["PPTM", "PPT", "POTX", "POTM", "POT", "PPSX", "PPSM", "PPS", "PPAM", "PowerPoint"]),
        ("PDF and fixed layout", ["fixed"], ["DjVu"]),
        ("Diagrams", [], ["Visio", "VSDX", "VSSX", "VSTX", "Binary Visio", "Binary Visio stencil", "Binary Visio template", "ODG/FODG"]),
        ("Project and Publisher", [], ["Project", "Publisher"]),
        ("Images", ["images"], []),
        ("Notes and books", ["books"], ["Compiled HTML Help"]),
        ("Notebooks", [], ["Jupyter Notebook", "Notebook source"]),
        ("Email", ["email"], []),
        ("Bibliography", ["bibliography"], []),
        ("Markup and text", ["markup"], []),
        ("Web and wiki", ["web"], []),
        ("Data", [], ["JSON", "JSON Lines", "XML", "YAML", "TOML"]),
        ("Other", ["other"], [])
    ];

    internal static string Create(string readmePath, FormatMapModel model) {
        string raw = File.ReadAllText(readmePath);
        string updated = Replace(Normalize(raw), Render(model));
        // Keep the README's own line endings, so regeneration only touches the block.
        return raw.Contains("\r\n", StringComparison.Ordinal) ? updated.Replace("\n", "\r\n") : updated;
    }

    internal static string Render(FormatMapModel model) {
        var placed = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int row = 0; row < Rows.Length; row++) {
            foreach (string format in Rows[row].Formats) {
                if (!placed.TryAdd(format, row)) throw new InvalidOperationException($"README format map places '{format}' twice.");
            }
        }

        var members = Rows.Select(static _ => new List<string>()).ToArray();
        foreach (FormatMapFormat format in model.Formats) {
            if (!placed.TryGetValue(format.Name, out int row)) {
                row = Array.FindIndex(Rows, candidate => candidate.WebsiteFamilies.Contains(format.Family, StringComparer.Ordinal));
                if (row < 0) row = Rows.Length - 1;
            }
            members[row].Add(format.Name);
        }

        var output = new StringBuilder();
        output.AppendLine(StartMarker)
            .Append(Count(model.Formats.Count)).Append(" formats and ").Append(Count(model.Routes.Count))
            .Append(" conversion routes. Every route runs from .NET; ")
            .Append(Count(Surface(model, "cli"))).Append(" run from the `officeimo` CLI, ")
            .Append(Count(Surface(model, "studio"))).Append(" in Studio, ")
            .Append(Count(Surface(model, "powershell"))).Append(" from PSWriteOffice and ")
            .Append(Count(Surface(model, "browser"))).AppendLine(" in the browser tools on officeimo.com.")
            .AppendLine()
            .AppendLine("| Family | Formats | Routes from |")
            .AppendLine("| --- | --- | ---: |");
        for (int row = 0; row < Rows.Length; row++) {
            if (members[row].Count == 0) continue;
            var formats = new HashSet<string>(members[row], StringComparer.Ordinal);
            int routes = model.Routes.Count(route => formats.Contains(route.Source));
            output.Append("| ").Append(Rows[row].Title)
                .Append(" | ").Append(string.Join(", ", members[row]))
                .Append(" | ").Append(Count(routes)).AppendLine(" |");
        }
        output.AppendLine()
            .AppendLine("Pick two formats on the [interactive format map](https://officeimo.com/) to see the route between them, or read the [conversion route catalog](Docs/Compatibility/generated/conversion-routes.md) for each route's package, API, support level and limits.")
            .AppendLine(EndMarker);
        return output.ToString();
    }

    internal static string Replace(string content, string block) {
        string[] lines = content.Split('\n');
        int start = Array.IndexOf(lines, StartMarker);
        int end = Array.IndexOf(lines, EndMarker);
        int starts = lines.Count(static line => line.Contains(StartMarker, StringComparison.Ordinal));
        int ends = lines.Count(static line => line.Contains(EndMarker, StringComparison.Ordinal));
        if (starts != 1 || ends != 1 || start < 0 || end < 0 || start >= end) {
            throw new InvalidDataException("The root README must contain exactly one format map block, with each marker on its own line.");
        }
        string before = string.Join("\n", lines.Take(start));
        string after = string.Join("\n", lines.Skip(end + 1));
        return before + "\n" + block.TrimEnd('\n') + "\n" + after;
    }

    private static int Surface(FormatMapModel model, string surface) =>
        model.Routes.Count(route => route.Surfaces.Contains(surface, StringComparer.Ordinal));

    private static string Count(int value) => value.ToString(CultureInfo.InvariantCulture);

    private static string Normalize(string value) => value.Replace("\r\n", "\n").Replace("\r", "\n");
}
