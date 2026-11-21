namespace OfficeIMO.Pdf.Testing;

/// <summary>Compact line diff of two structure files for exception messages.</summary>
internal static class PdfSnapshotDiff {
    private const int MaximumLines = 40;
    private const int Context = 2;

    public static string Describe(string approved, string received) {
        string[] expected = Normalize(approved).Split('\n');
        string[] actual = Normalize(received).Split('\n');
        int prefix = 0;
        while (prefix < expected.Length && prefix < actual.Length && expected[prefix] == actual[prefix]) prefix++;
        int suffix = 0;
        while (suffix < expected.Length - prefix && suffix < actual.Length - prefix
            && expected[expected.Length - 1 - suffix] == actual[actual.Length - 1 - suffix]) suffix++;

        var lines = new List<string>();
        for (int index = Math.Max(0, prefix - Context); index < prefix; index++) lines.Add("  " + expected[index]);
        for (int index = prefix; index < expected.Length - suffix; index++) lines.Add("- " + expected[index]);
        for (int index = prefix; index < actual.Length - suffix; index++) lines.Add("+ " + actual[index]);
        for (int index = expected.Length - suffix; index < Math.Min(expected.Length, expected.Length - suffix + Context); index++) lines.Add("  " + expected[index]);

        var text = new StringBuilder();
        text.Append("@@ approved line ").Append(prefix + 1).Append(" @@\n");
        foreach (string line in lines.Take(MaximumLines)) text.Append(line).Append('\n');
        if (lines.Count > MaximumLines) text.Append("  ... ").Append(lines.Count - MaximumLines).Append(" more lines; compare the structure files.\n");
        return text.ToString();
    }

    public static string Normalize(string text) => text.Replace("\r\n", "\n");
}
