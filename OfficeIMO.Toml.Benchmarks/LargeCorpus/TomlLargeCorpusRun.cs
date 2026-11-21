using System.Diagnostics;
using System.Text;

namespace OfficeIMO.Toml.Benchmarks.LargeCorpus;

/// <summary>
/// <c>--large-corpus</c>: every pinned file is loaded, compared with Tomlyn's data, saved back byte for byte, read through
/// the stream reader with the same tokens as in memory, and edited with a fixed script whose result is reloaded and
/// compared with the edited model. Writes <c>large-corpus.md</c> and <c>large-corpus-failures.tsv</c>.
/// </summary>
internal static class TomlLargeCorpusRun {
    private const string AddedKey = "officeimo-added";
    private const string AddedTable = "officeimo-table";

    private sealed class Totals {
        public int Files;
        public long Bytes;
        public int Loaded;
        public int Rejected;
        public int TomlynAgrees;
        public int TomlynRejectsOnly;
        public int TomlynAcceptsOnly;
        public int DataDiffers;
        public int RoundTrips;
        public int StreamMatches;
        public int Edited;
        public int EditsVerified;
        public int MaxAddedLines;
        public int MaxRemovedLines;
        public TimeSpan LoadTime;
        public TimeSpan TomlynTime;
        public readonly List<string> Failures = new();
        public readonly SortedDictionary<string, (int Files, long Bytes)> Categories = new(StringComparer.Ordinal);
    }

    public static int Run(string[] args) {
        string root = TomlLargeCorpus.ResolveRoot(TomlLargeCorpus.Option(args, "--cache"));
        string output = TomlLargeCorpus.Option(args, "--output-dir") ?? Path.Combine(root, "reports");
        List<TomlCorpusFile> files = TomlLargeCorpus.Files(root);
        if (TomlLargeCorpus.Option(args, "--file") is string only) files = files.Where(f => f.Name == only || f.Path == only).ToList();
        var totals = new Totals();
        foreach (TomlCorpusFile file in files) {
            totals.Files++;
            totals.Bytes += file.Bytes;
            totals.Categories[file.Category] = (totals.Categories.GetValueOrDefault(file.Category).Files + 1, totals.Categories.GetValueOrDefault(file.Category).Bytes + file.Bytes);
            try {
                Check(file, file.Read(), totals);
            } catch (Exception ex) {
                totals.Failures.Add(file.Name + "\tcrash\t" + ex.GetType().Name + ": " + ex.Message);
            }
        }
        Directory.CreateDirectory(output);
        File.WriteAllText(Path.Combine(output, "large-corpus.md"), Report(totals));
        File.WriteAllLines(Path.Combine(output, "large-corpus-failures.tsv"), totals.Failures);
        Console.Write(Report(totals));
        // Rejections and differences from Tomlyn are findings to review; only crashes, round-trip, stream-reader and edit
        // failures are OfficeIMO defects.
        int defects = totals.Failures.Count(f => f.Contains("\tcrash\t") || f.Contains("\tround-trip\t") || f.Contains("\tstream reader\t") || f.Contains("\tedit\t"));
        return defects == 0 ? 0 : 1;
    }

    private static void Check(TomlCorpusFile file, byte[] bytes, Totals totals) {
        string text = DecodeText(bytes);

        TomlDocument? document = null;
        string? error = null;
        var watch = Stopwatch.StartNew();
        try {
            document = TomlDocument.Load(new MemoryStream(bytes));
        } catch (TomlException ex) {
            error = ex.Code + " at " + ex.Position + ": " + ex.Message;
        }
        totals.LoadTime += watch.Elapsed;

        object? tomlyn = null;
        string? tomlynError = null;
        watch.Restart();
        try {
            tomlyn = Tomlyn.TomlSerializer.Deserialize<Tomlyn.Model.TomlTable>(text);
        } catch (Exception ex) {
            tomlynError = ex.Message.Split('\n')[0];
        }
        totals.TomlynTime += watch.Elapsed;

        if (document is null) {
            totals.Rejected++;
            if (tomlyn is null) totals.TomlynAgrees++;
            else totals.TomlynAcceptsOnly++;
            totals.Failures.Add(file.Name + "\t" + (tomlyn is null ? "rejected by both" : "rejected (Tomlyn accepts)") + "\t" + error);
            return;
        }
        using (document) {
            totals.Loaded++;
            string data = TomlCorpusData.Of(document.Root);
            if (tomlyn is null) {
                totals.TomlynRejectsOnly++;
            } else if (data == TomlCorpusData.OfTomlyn(tomlyn)) {
                totals.TomlynAgrees++;
            } else {
                totals.DataDiffers++;
                totals.Failures.Add(file.Name + "\tdata differs from Tomlyn\t" + FirstDifference(data, TomlCorpusData.OfTomlyn(tomlyn)));
            }

            var saved = new MemoryStream();
            document.Save(saved);
            if (saved.ToArray().AsSpan().SequenceEqual(bytes)) totals.RoundTrips++;
            else totals.Failures.Add(file.Name + "\tround-trip\tsaved bytes differ");

            if (StreamTrace(bytes) == MemoryTrace(text)) totals.StreamMatches++;
            else totals.Failures.Add(file.Name + "\tstream reader\ttokens differ from in-memory reading");
        }

        string? editFailure = Edit(bytes, totals);
        if (editFailure != null) totals.Failures.Add(file.Name + "\tedit\t" + editFailure);
    }

    /// <summary>
    /// Sets the first scalar in a table to a new value of its type, adds a key at the root, removes the last value of the
    /// last table that has one, and adds a table with a key. The saved text must reload to the edited model and differ
    /// from the original in only a few lines.
    /// </summary>
    private static string? Edit(byte[] bytes, Totals totals) {
        using TomlDocument document = TomlDocument.Load(new MemoryStream(bytes));
        if (document.Root.ContainsKey(AddedKey) || document.Root.ContainsKey(AddedTable)) return null;
        totals.Edited++;
        var tables = new List<TomlTable>();
        CollectTables(document.Root, tables);
        TomlTable? setOwner = null;
        int setLines = 0;
        string? setKey = null;
        foreach (TomlTable table in tables) {
            foreach (KeyValuePair<string, TomlNode> entry in table) {
                if (entry.Value is TomlValue value && value.Kind != TomlNodeKind.DateTime) {
                    setOwner = table;
                    setKey = entry.Key;
                    break;
                }
            }
            if (setOwner != null) break;
        }
        if (setOwner != null) {
            object current = ((TomlValue)setOwner[setKey!]).Value;
            object next = current switch {
                string s => s + "-edited",
                long l => l == long.MaxValue ? l - 1 : l + 1,
                double d => double.IsFinite(d) ? d + 0.5 : 1.5,
                bool b => !b,
                _ => current
            };
            TomlNode target = setOwner[setKey!];
            setLines = target.End.Line - target.Start.Line + 1;
            setOwner.Set(setKey!, next);
        }
        int commentLines = 0, removedValueLines = 1;
        for (int i = tables.Count - 1; i >= 0; i--) {
            TomlTable table = tables[i];
            if (ReferenceEquals(table, setOwner)) continue;
            KeyValuePair<string, TomlNode> last = table.LastOrDefault(e => e.Value is TomlValue);
            if (last.Key is null) continue;
            commentLines = last.Value.LeadingComments.Count;
            removedValueLines = last.Value.End.Line - last.Value.Start.Line + 1;
            table.Remove(last.Key);
            break;
        }
        document.Root.Add(AddedKey, "added value");
        TomlTable added = document.Root.AddTable(AddedTable);
        added.Add("enabled", true);

        string expected = TomlCorpusData.Of(document.Root);
        string saved = document.ToString();
        TomlDocument reloaded;
        try {
            reloaded = TomlDocument.Parse(saved);
        } catch (TomlException ex) {
            return "edited text does not load: " + ex.Code + " at " + ex.Position + ": " + ex.Message;
        }
        using (reloaded) {
            if (TomlCorpusData.Of(reloaded.Root) != expected) return "edited text loads to different data: " + FirstDifference(expected, TomlCorpusData.Of(reloaded.Root));
        }
        (int addedLines, int removedLines) = LineDifference(DecodeText(bytes), saved);
        totals.MaxAddedLines = Math.Max(totals.MaxAddedLines, addedLines);
        totals.MaxRemovedLines = Math.Max(totals.MaxRemovedLines, removedLines);
        // Set: the value's lines out and in. Remove: its lines and comments. Added key and table: a few lines and blanks.
        if (removedLines > setLines + removedValueLines + commentLines || addedLines > setLines + 6) return "the edit changed " + removedLines + " removed and " + addedLines + " added lines";
        totals.EditsVerified++;
        return null;
    }

    private static void CollectTables(TomlTable table, List<TomlTable> tables) {
        if (table.TableKind == TomlTableKind.Inline) return;
        tables.Add(table);
        foreach (KeyValuePair<string, TomlNode> entry in table) {
            if (entry.Value is TomlTable child) CollectTables(child, tables);
            else if (entry.Value is TomlTableArray array) {
                foreach (TomlTable item in array) CollectTables(item, tables);
            }
        }
    }

    /// <summary>Counts lines only in the original and only in the edited text, as multisets.</summary>
    private static (int Added, int Removed) LineDifference(string original, string edited) {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string line in original.Split('\n')) counts[line] = counts.GetValueOrDefault(line) + 1;
        int added = 0;
        foreach (string line in edited.Split('\n')) {
            int count = counts.GetValueOrDefault(line);
            if (count > 0) counts[line] = count - 1;
            else added++;
        }
        return (added, counts.Values.Sum());
    }

    private static string MemoryTrace(string text) => Trace(new TomlReader(text, new TomlReaderOptions { IncludeComments = true }));

    private static string StreamTrace(byte[] bytes) => Trace(new TomlReader(new MemoryStream(bytes), new TomlReaderOptions { IncludeComments = true }));

    private static string Trace(TomlReader reader) {
        var trace = new StringBuilder();
        using (reader) {
            try {
                while (reader.Read()) trace.Append((int)reader.TokenType).Append(' ').Append(reader.Start.Offset).Append(' ').Append(reader.End).Append('\n');
            } catch (TomlException ex) {
                trace.Append(ex.Code).Append(ex.Position);
            }
        }
        return trace.ToString();
    }

    internal static string DecodeText(byte[] bytes) {
        int offset = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
        return Encoding.UTF8.GetString(bytes, offset, bytes.Length - offset);
    }

    private static string FirstDifference(string expected, string actual) {
        int i = 0;
        while (i < expected.Length && i < actual.Length && expected[i] == actual[i]) i++;
        int start = Math.Max(0, i - 60);
        return "at " + i + ": '" + Clip(expected, start) + "' vs '" + Clip(actual, start) + "'";
    }

    private static string Clip(string text, int start) => start >= text.Length ? string.Empty : text.Substring(start, Math.Min(140, text.Length - start)).Replace('\n', ' ').Replace('\t', ' ');

    private static string Report(Totals t) {
        var report = new StringBuilder();
        report.AppendLine("# TOML large corpus");
        report.AppendLine();
        report.AppendLine("| Measure | Files |");
        report.AppendLine("| --- | ---: |");
        report.AppendLine("| Pinned files | " + TomlLargeCorpus.Format(t.Files) + " (" + TomlLargeCorpus.Format(t.Bytes) + " bytes) |");
        report.AppendLine("| Loaded | " + TomlLargeCorpus.Format(t.Loaded) + " |");
        report.AppendLine("| Rejected | " + TomlLargeCorpus.Format(t.Rejected) + " |");
        report.AppendLine("| Same acceptance and data as Tomlyn | " + TomlLargeCorpus.Format(t.TomlynAgrees) + " |");
        report.AppendLine("| Loaded; Tomlyn rejects | " + TomlLargeCorpus.Format(t.TomlynRejectsOnly) + " |");
        report.AppendLine("| Rejected; Tomlyn accepts | " + TomlLargeCorpus.Format(t.TomlynAcceptsOnly) + " |");
        report.AppendLine("| Data differs from Tomlyn | " + TomlLargeCorpus.Format(t.DataDiffers) + " |");
        report.AppendLine("| Saved back byte for byte | " + TomlLargeCorpus.Format(t.RoundTrips) + " |");
        report.AppendLine("| Stream reader tokens equal to in-memory | " + TomlLargeCorpus.Format(t.StreamMatches) + " |");
        report.AppendLine("| Scripted edits reload to the edited data with a minimal diff | " + TomlLargeCorpus.Format(t.EditsVerified) + " of " + TomlLargeCorpus.Format(t.Edited) + " |");
        report.AppendLine();
        report.AppendLine("Largest edit diff: " + t.MaxRemovedLines + " lines removed, " + t.MaxAddedLines + " added. Load time " + t.LoadTime.TotalSeconds.ToString("F2") + " s (OfficeIMO) and " + t.TomlynTime.TotalSeconds.ToString("F2") + " s (Tomlyn), single-threaded, including first-call JIT.");
        report.AppendLine();
        report.AppendLine("| Category | Files | Bytes |");
        report.AppendLine("| --- | ---: | ---: |");
        foreach (KeyValuePair<string, (int Files, long Bytes)> category in t.Categories) {
            report.AppendLine("| " + category.Key + " | " + TomlLargeCorpus.Format(category.Value.Files) + " | " + TomlLargeCorpus.Format(category.Value.Bytes) + " |");
        }
        report.AppendLine();
        report.AppendLine("Findings: " + t.Failures.Count + " (see large-corpus-failures.tsv).");
        return report.ToString();
    }
}
