using OfficeIMO.Tool.Commands.Convert;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeIMO.Tool;

/// <summary>Writes the machine-readable command catalog printed by <c>officeimo help --format json</c>.</summary>
/// <remarks>Usage lines, subcommand names and option names are read from each area's help text, so the catalog and <c>--help</c> cannot disagree.</remarks>
internal static class OfficeImoToolCommandCatalog {
    internal const string Schema = "officeimo.tool.commands.v1";

    private static readonly Regex OptionPattern = new(
        @"(?<![\w-])(?<name>--[a-z0-9][a-z0-9-]*)(?:\s+(?<value><[^>]+>|[^\s\[\]<>|-][^\s\[\]]*))?",
        RegexOptions.CultureInvariant);

    // A bracketed note with a semicolon, such as "[print-plan options; --paper defaults to ...]", is prose rather than syntax.
    private static readonly Regex ProseNotePattern = new(@"\[[^\]]*;[^\]]*\]", RegexOptions.CultureInvariant);

    private static readonly Regex DescribedOptionPattern = new(
        @"^  (?<name>--[a-z0-9][a-z0-9-]*)(?: (?<value><[^>]+>|\S+))?\s{2,}(?<description>\S.*)$",
        RegexOptions.CultureInvariant);

    internal static void Write(Stream output, string version) {
        using var writer = new Utf8JsonWriter(output, new JsonWriterOptions { Indented = true });
        writer.WriteStartObject();
        writer.WriteString("schema", Schema);
        writer.WriteStartObject("tool");
        writer.WriteString("command", "officeimo");
        writer.WriteString("package", "OfficeIMO.Tool");
        writer.WriteString("version", version);
        writer.WriteEndObject();

        writer.WriteStartArray("groups");
        foreach (OfficeImoToolCommandGroup group in OfficeImoToolCommands.Groups) WriteGroup(writer, group);
        writer.WriteEndArray();

        writer.WriteStartArray("exitCodes");
        foreach (OfficeImoToolExitCode code in Enum.GetValues<OfficeImoToolExitCode>()) {
            writer.WriteStartObject();
            writer.WriteNumber("code", (int)code);
            writer.WriteString("name", code.ToString());
            writer.WriteString("meaning", OfficeImoToolExitCodes.Describe(code));
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    private static void WriteGroup(Utf8JsonWriter writer, OfficeImoToolCommandGroup group) {
        IReadOnlyList<UsageEntry> entries = ReadUsage(group);
        var subcommands = new List<(string Name, List<string> Usage)>();
        foreach (UsageEntry entry in entries.Where(static entry => entry.Name.Length > 0)) {
            int index = subcommands.FindIndex(existing => existing.Name == entry.Name);
            if (index < 0) subcommands.Add((entry.Name, [entry.Line]));
            else subcommands[index].Usage.Add(entry.Line);
        }

        writer.WriteStartObject();
        writer.WriteString("id", group.Id);
        writer.WriteString("title", group.Title);
        writer.WriteString("summary", group.Summary);
        WriteStrings(writer, "aliases", group.Aliases);
        if (group.ShortcutFor is null) writer.WriteNull("shortcutFor");
        else writer.WriteString("shortcutFor", group.ShortcutFor);
        WriteStrings(writer, "usage", group.ShortcutFor is null
            ? entries.Where(static entry => entry.Name.Length == 0).Select(static entry => entry.Line).DefaultIfEmpty(TopLevelUsage(group.Id, group.Synopsis))
            : new[] { group.Id }.Concat(group.Aliases).Select(name => TopLevelUsage(name, group.Synopsis)));

        writer.WriteStartArray("subcommands");
        foreach ((string name, List<string> usage) in subcommands) {
            if (!group.Subcommands.TryGetValue(name, out string? summary)) {
                throw new InvalidOperationException("The '" + group.Id + " " + name + "' command has no catalog summary.");
            }
            writer.WriteStartObject();
            writer.WriteString("name", name);
            WriteStrings(writer, "names", ExpandAlternatives(name));
            writer.WriteString("summary", summary);
            WriteStrings(writer, "usage", usage);
            writer.WriteStartArray("options");
            foreach ((string optionName, string? value) in usage.SelectMany(ReadOptions).DistinctBy(static option => option.Name)) {
                writer.WriteStartObject();
                writer.WriteString("name", optionName);
                if (value is null) writer.WriteNull("value");
                else writer.WriteString("value", value);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        writer.WriteEndArray();

        writer.WriteStartArray("keyOptions");
        foreach ((string name, string? value, string description) in KeyOptions(group, entries)) {
            writer.WriteStartObject();
            writer.WriteString("name", name);
            if (value is null) writer.WriteNull("value");
            else writer.WriteString("value", value);
            writer.WriteString("description", description);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();

        WriteStrings(writer, "inputFormats", NormalizeFormats(group.InputFormats()));
        WriteStrings(writer, "outputFormats", NormalizeFormats(group.OutputFormats()));
        WriteStrings(writer, "resultFormats", group.ResultFormats);

        if (group.Id == "convert") {
            writer.WriteStartArray("routes");
            foreach (var route in ConvertRoutes.All()) {
                writer.WriteStartObject();
                writer.WriteString("id", route.Id);
                writer.WriteString("source", route.Source);
                writer.WriteString("target", route.Target);
                WriteStrings(writer, "sourceExtensions", route.SourceExtensions);
                writer.WriteString("targetExtension", route.TargetExtension);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
        writer.WriteEndObject();
    }

    internal static string TopLevelUsage(string name, string synopsis) => "officeimo " + name + " " + synopsis;

    private static IEnumerable<(string Name, string? Value, string Description)> KeyOptions(OfficeImoToolCommandGroup group, IReadOnlyList<UsageEntry> entries) {
        // Values come from usage lines first, then from option lists such as the invoice "Target options" section.
        IEnumerable<string> optionLines = group.Help
            .SelectMany(static help => help.Split('\n'))
            .Where(static line => line.StartsWith("  --", StringComparison.Ordinal));
        Dictionary<string, string?> values = entries
            .Select(static entry => entry.Line)
            .Concat(optionLines)
            .SelectMany(ReadOptions)
            .DistinctBy(static option => option.Name)
            .ToDictionary(static option => option.Name, static option => option.Value, StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeImoToolKeyOption option in group.KeyOptions) {
            if (!seen.Add(option.Name)) continue;
            values.TryGetValue(option.Name, out string? value);
            yield return (option.Name, option.Value ?? value, option.Description);
        }
        // Options the help text itself describes, such as the markup "Options:" section.
        foreach (string line in group.Help.SelectMany(static help => help.Split('\n'))) {
            Match match = DescribedOptionPattern.Match(line.TrimEnd('\r'));
            if (!match.Success || !seen.Add(match.Groups["name"].Value)) continue;
            yield return (
                match.Groups["name"].Value,
                match.Groups["value"].Success ? match.Groups["value"].Value : null,
                match.Groups["description"].Value.Trim());
        }
    }

    internal static IReadOnlyList<UsageEntry> ReadUsage(OfficeImoToolCommandGroup group) {
        string prefix = "officeimo " + group.Id;
        var entries = new List<UsageEntry>();
        foreach (string help in group.Help) {
            StringBuilder? current = null;
            foreach (string rawLine in help.Split('\n')) {
                string line = rawLine.TrimEnd('\r');
                string trimmed = line.Trim();
                bool starts = line.StartsWith("  ", StringComparison.Ordinal)
                    && (trimmed == prefix || trimmed.StartsWith(prefix + " ", StringComparison.Ordinal));
                if (starts) {
                    Flush(current, prefix, entries);
                    current = new StringBuilder(trimmed);
                } else if (current is not null && line.StartsWith("   ", StringComparison.Ordinal) && trimmed.Length > 0) {
                    current.Append(' ').Append(trimmed);
                } else {
                    Flush(current, prefix, entries);
                    current = null;
                }
            }
            Flush(current, prefix, entries);
        }
        return entries;
    }

    private static void Flush(StringBuilder? current, string prefix, List<UsageEntry> entries) {
        if (current is null) return;
        string line = Regex.Replace(current.ToString(), @"\s+", " ");
        string[] tokens = line[prefix.Length..].Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var name = new List<string>();
        int index = 0;
        for (; index < tokens.Length && tokens[index][0] is not ('<' or '[' or '-'); index++) name.Add(tokens[index]);
        // "officeimo pdf redact <command>" points at a more specific help block.
        if (name.Count > 0 && index < tokens.Length && tokens[index] == "<command>") return;
        entries.Add(new UsageEntry(string.Join(' ', name), line));
    }

    internal static IEnumerable<(string Name, string? Value)> ReadOptions(string usageLine) {
        foreach (Match match in OptionPattern.Matches(ProseNotePattern.Replace(usageLine, string.Empty))) {
            yield return (match.Groups["name"].Value, match.Groups["value"].Success ? match.Groups["value"].Value : null);
        }
    }

    private static IEnumerable<string> ExpandAlternatives(string name) {
        int split = name.LastIndexOf(' ');
        string head = split < 0 ? string.Empty : name[..(split + 1)];
        return name[(split + 1)..].Split('|').Select(alternative => head + alternative);
    }

    private static IEnumerable<string> NormalizeFormats(IEnumerable<string> formats) => formats
        .Select(static format => format.TrimStart('.').ToLowerInvariant())
        .Where(static format => format.Length > 0)
        .Distinct(StringComparer.Ordinal)
        .Order(StringComparer.Ordinal);

    private static void WriteStrings(Utf8JsonWriter writer, string name, IEnumerable<string> values) {
        writer.WriteStartArray(name);
        foreach (string value in values) writer.WriteStringValue(value);
        writer.WriteEndArray();
    }

    internal readonly record struct UsageEntry(string Name, string Line);
}
