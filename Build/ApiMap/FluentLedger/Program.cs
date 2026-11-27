using System.Diagnostics;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using OfficeIMO.Build.ApiMap;
using OfficeIMO.Build.ApiMap.Fluent;

// Usage: FluentLedger <repository-root> [--config <path>] [--output <directory>] [--framework <tfm> ...]
if (args.Length < 1) {
    Console.Error.WriteLine("Usage: FluentLedger <repository-root> [--config <path>] [--output <directory>] [--framework <tfm> ...]");
    return 2;
}
string root = Path.GetFullPath(args[0]);
string configPath = Path.Combine(root, "Build", "ApiMap", "FluentLedger", "fluent-ledger.json");
string output = Path.Combine(root, "Build", "ApiMap", "fluent");
var frameworks = new List<string>();
for (int i = 1; i < args.Length; i++) {
    switch (args[i]) {
        case "--config": configPath = Path.GetFullPath(args[++i]); break;
        case "--output": output = Path.GetFullPath(args[++i]); break;
        case "--framework": frameworks.Add(args[++i]); break;
        default: Console.Error.WriteLine("Unknown argument: " + args[i]); return 2;
    }
}
var config = LedgerConfig.Load(configPath);
if (frameworks.Count == 0) frameworks.AddRange(config.Frameworks);

var start = new ProcessStartInfo("git") { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true };
foreach (string argument in new[] { "ls-files", "--cached", "--others", "--exclude-standard" }) start.ArgumentList.Add(argument);
using (var process = Process.Start(start) ?? throw new InvalidOperationException("Could not enumerate repository inputs.")) {
    string input = process.StandardOutput.ReadToEnd();
    process.WaitForExit();
    if (process.ExitCode != 0) throw new InvalidOperationException(process.StandardError.ReadToEnd());
    var files = input.Split('\n', StringSplitOptions.RemoveEmptyEntries).Distinct(StringComparer.Ordinal).ToArray();
    var source = new SourceLocationIndex(root, files, config.Packages.ToHashSet(StringComparer.Ordinal));

    // One row per member and receiver; a member present in several framework builds lists each of them.
    var merged = new Dictionary<(string Id, string Receiver), LedgerRow>();
    var counts = new List<string>();
    foreach (string framework in frameworks) {
        var ledger = new Ledger(config, source);
        var rows = ledger.Build(new MetadataCatalog(root, config.Packages, framework), framework);
        counts.Add($"{framework}: {ledger.PublicTypes} public types, {ledger.Receivers} authoring receivers, {rows.Count} members");
        foreach (var row in rows) {
            var key = (row.Id, row.Receiver);
            merged[key] = merged.TryGetValue(key, out var existing)
                ? existing with { Frameworks = existing.Frameworks!.Concat(row.Frameworks!).Distinct().ToArray() }
                : row;
        }
    }
    LedgerOutput.Write(output, config, frameworks, merged.Values.ToArray(), counts);
    foreach (string line in counts) Console.WriteLine(line);
}
return 0;

internal static class LedgerOutput {
    private static readonly JsonSerializerOptions Json = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    internal static void Write(string output, LedgerConfig config, IReadOnlyList<string> frameworks, IReadOnlyList<LedgerRow> rows, IReadOnlyList<string> counts) {
        Directory.CreateDirectory(output);
        foreach (string stale in Directory.GetFiles(output, "*.json")) File.Delete(stale);
        var byPackage = rows.GroupBy(row => row.Package!).OrderBy(group => group.Key, StringComparer.Ordinal).ToArray();
        foreach (var package in byPackage) {
            // One row per line keeps regenerated diffs reviewable.
            var builder = new StringBuilder();
            builder.Append("{\n  \"package\": ").Append(JsonSerializer.Serialize(package.Key, Json))
                .Append(",\n  \"generatedBy\": \"Build/ApiMap/FluentLedger\",\n  \"frameworks\": ").Append(JsonSerializer.Serialize(frameworks, Json))
                .Append(",\n  \"rows\": [\n");
            var ordered = Order(package).Select(row => (row.Frameworks?.Count == frameworks.Count ? row with { Frameworks = null } : row) with { Package = null }).ToArray();
            for (int i = 0; i < ordered.Length; i++)
                builder.Append("    ").Append(JsonSerializer.Serialize(ordered[i], Json)).Append(i + 1 < ordered.Length ? ",\n" : "\n");
            builder.Append("  ]\n}\n");
            File.WriteAllText(Path.Combine(output, package.Key + ".json"), builder.ToString());
        }
        File.WriteAllText(Path.Combine(output, "README.md"), Summary(config, frameworks, rows, byPackage, counts));
    }

    private static IEnumerable<LedgerRow> Order(IEnumerable<LedgerRow> rows) => rows
        .OrderBy(row => row.Receiver, StringComparer.Ordinal).ThenBy(row => row.Member, StringComparer.Ordinal).ThenBy(row => row.Signature, StringComparer.Ordinal);

    private static readonly string[] Columns = [
        Ledger.Treatments.Rename, Ledger.Treatments.Generate, Ledger.Treatments.GenerateNested, Ledger.Treatments.GenerateConfigure,
        Ledger.Treatments.KeepAction, Ledger.Treatments.KeepCopy, Ledger.Treatments.KeepAdd, Ledger.Treatments.KeepLifecycle,
        Ledger.Treatments.KeepConversion, Ledger.Treatments.KeepSelector, Ledger.Treatments.KeepProperty, Ledger.Treatments.KeepFluent, Ledger.Treatments.Keep,
        Ledger.Treatments.Retire, Ledger.Treatments.Decision ];

    private static string Summary(LedgerConfig config, IReadOnlyList<string> frameworks, IReadOnlyList<LedgerRow> rows, IReadOnlyList<IGrouping<string, LedgerRow>> byPackage, IReadOnlyList<string> counts) {
        var text = new StringBuilder();
        text.AppendLine("# 4.0 fluent symbol ledger");
        text.AppendLine();
        text.AppendLine("<!-- Generated by Build/ApiMap/FluentLedger. Do not edit; change the rules or fluent-ledger.json and regenerate. -->");
        text.AppendLine();
        text.AppendLine($"One row per exported member and overload of every authoring receiver, with its proposed 4.0 form under the [#3344]({config.Issue}) maps. ");
        text.AppendLine($"Built from the Release assemblies for {string.Join(", ", frameworks)}. Treatments and flags are described in [the ledger tool's README](../FluentLedger/README.md).");
        text.AppendLine();
        foreach (string line in counts) text.AppendLine("- " + line);
        text.AppendLine();
        text.AppendLine("## Treatments by package");
        text.AppendLine();
        text.AppendLine("| Package | " + string.Join(" | ", Columns) + " | total |");
        text.AppendLine("| --- |" + string.Concat(Enumerable.Repeat(" ---: |", Columns.Length + 1)));
        foreach (var package in byPackage) {
            var tally = package.GroupBy(row => row.Treatment).ToDictionary(group => group.Key, group => group.Count());
            text.AppendLine($"| [{package.Key}]({package.Key}.json) | " + string.Join(" | ", Columns.Select(column => tally.TryGetValue(column, out int count) ? count.ToString() : "")) + $" | {package.Count()} |");
        }
        var all = rows.GroupBy(row => row.Treatment).ToDictionary(group => group.Key, group => group.Count());
        text.AppendLine("| **All** | " + string.Join(" | ", Columns.Select(column => all.TryGetValue(column, out int count) ? count.ToString() : "")) + $" | {rows.Count} |");
        text.AppendLine();

        var decisions = rows.Where(row => row.Treatment == Ledger.Treatments.Decision).ToArray();
        text.AppendLine("## Decisions by reason");
        text.AppendLine();
        text.AppendLine("Every `decision` row carries one of these codes in `flags`.");
        text.AppendLine();
        text.AppendLine("| Code | Rows | Receivers | Examples |");
        text.AppendLine("| --- | ---: | ---: | --- |");
        foreach (var group in decisions.GroupBy(DecisionCode).OrderByDescending(group => group.Count()).ThenBy(group => group.Key, StringComparer.Ordinal)) {
            var examples = Order(group).Select(row => $"`{Short(row.Receiver)}.{row.Member}`").Distinct().Take(6);
            text.AppendLine($"| `{group.Key}` | {group.Count()} | {group.Select(row => row.Receiver).Distinct().Count()} | {string.Join(", ", examples)} |");
        }
        text.AppendLine();

        text.AppendLine("## Rename families");
        text.AppendLine();
        text.AppendLine("Set-style configuration methods that become property-named extensions, by family.");
        text.AppendLine();
        text.AppendLine("| Current | 4.0 | Overloads | Receivers |");
        text.AppendLine("| --- | --- | ---: | ---: |");
        foreach (var family in rows.Where(row => row.Treatment == Ledger.Treatments.Rename && row.Kind != "property")
                     .GroupBy(row => (row.Member, Target: row.Proposed![..row.Proposed!.IndexOf('(')]))
                     .OrderByDescending(group => group.Count()).ThenBy(group => group.Key.Member, StringComparer.Ordinal))
            text.AppendLine($"| `{family.Key.Member}` | `{family.Key.Target}` | {family.Count()} | {family.Select(row => row.Receiver).Distinct().Count()} |");
        return text.ToString();
    }

    private static readonly string[] DecisionCodes = [
        "name-collision", "generic-set", "getter-wraps", "returns-other", "bare-set", "use-prefix", "static-set", "set-returns-copy", "with-mutates", "struct-property", "public-field" ];

    private static string DecisionCode(LedgerRow row) =>
        (row.Flags ?? []).FirstOrDefault(flag => DecisionCodes.Contains(flag)) ?? (row.Flags ?? []).FirstOrDefault(flag => !flag.StartsWith("struct-", StringComparison.Ordinal) && flag is not ("void-setter" or "no-property" or "interface-member" or "multi-value")) ?? "other";

    private static string Short(string receiver) => receiver[(receiver.LastIndexOf('.') + 1)..];
}
