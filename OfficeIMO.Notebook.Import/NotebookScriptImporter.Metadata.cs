using System.Text;
using System.Text.Json;
using OfficeIMO.Yaml;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    private readonly struct PercentHeader(string? language, bool declaresPercent, bool commentMagics) {
        internal string? Language => language;
        internal bool DeclaresPercent => declaresPercent;
        internal bool CommentMagics => commentMagics;
    }

    private sealed class PercentMarker(NotebookCellKind kind, string? id, string? language, Dictionary<string, JsonElement> metadata) {
        internal NotebookCellKind Kind => kind;
        internal string? Id => id;
        internal string? Language => language;
        internal Dictionary<string, JsonElement> Metadata => metadata;
    }

    private static PercentHeader ReadPercentHeader(StringReader reader, NotebookImportContext context, ref string? line, ref int lineNumber) {
        var yaml = new StringBuilder();
        while ((line = reader.ReadLine()) != null) {
            context.CancellationToken.ThrowIfCancellationRequested();
            lineNumber++;
            ReadOnlySpan<char> raw = line.AsSpan();
            if (raw.TrimEnd().SequenceEqual("# ---".AsSpan())) {
                line = reader.ReadLine();
                lineNumber++;
                if (yaml.Length == 0) return new PercentHeader(null, false, true);
                using JsonDocument parsed = ParsePercentYaml(yaml.ToString(), context);
                if (parsed.RootElement.ValueKind != JsonValueKind.Object) {
                    throw new FormatException("The Jupytext header must contain one YAML mapping.");
                }

                JsonElement jupyter = ReadObjectProperty(parsed.RootElement, "jupyter");
                if (jupyter.ValueKind == JsonValueKind.Undefined) jupyter = parsed.RootElement;
                if (jupyter.ValueKind != JsonValueKind.Object) throw new FormatException("The Jupytext jupyter header must be a metadata mapping.");
                foreach (JsonProperty property in jupyter.EnumerateObject()) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    context.Document.SetMetadata(property.Name, property.Value);
                }
                foreach (JsonProperty property in parsed.RootElement.EnumerateObject()) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    if (property.Name != "jupyter" && !context.Document.Metadata.ContainsKey(property.Name)) {
                        context.Document.SetMetadata(property.Name, property.Value);
                    }
                }

                JsonElement jupytext = ReadObjectProperty(jupyter, "jupytext");
                JsonElement format = ReadObjectProperty(jupytext, "text_representation");
                string? formatName = ReadStringProperty(format, "format_name");
                if (formatName != null && formatName != "percent") throw new FormatException("The Jupytext header declares a different script format.");
                JsonElement kernel = ReadObjectProperty(jupyter, "kernelspec");
                string? language = ReadStringProperty(kernel, "language")
                    ?? ReadStringProperty(ReadObjectProperty(jupyter, "language_info"), "name")
                    ?? ReadStringProperty(jupytext, "main_language")
                    ?? ReadStringProperty(kernel, "name");
                return new PercentHeader(language, formatName == "percent",
                    ReadObjectProperty(jupytext, "comment_magics").ValueKind != JsonValueKind.False);
            }
            if (raw.IsEmpty || raw[0] != '#') throw new FormatException("The Jupytext header is missing its closing '# ---' marker.");
            yaml.Append(RemoveComment(raw).ToString()).Append('\n');
        }

        throw new FormatException("The Jupytext header is missing its closing '# ---' marker.");
    }

    private static JsonDocument ParsePercentYaml(string source, NotebookImportContext context) {
        var limits = YamlLoadOptions.UntrustedDefaults;
        limits.MaxDepth = Math.Min(limits.MaxDepth, context.Limits.MaxJsonDepth);
        limits.MaxAliasExpansion = 1;
        limits.Strict = true;
        try {
            using var yaml = YamlDocument.Parse(source, limits, context.CancellationToken);
            if (yaml.Documents.Count > 1) throw new FormatException("A Jupytext YAML value must contain a single document.");
            if (yaml.Root != null) {
                foreach (YamlNode node in yaml.Root.Descendants()) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    if (node is YamlAlias) throw new FormatException("Jupytext metadata aliases are outside the supported profile.");
                }
            }
            return JsonDocument.Parse(yaml.ToJsonData().ToString(), JsonOptions(context));
        } catch (YamlLimitExceededException error) {
            throw new InvalidDataException("The Jupytext metadata exceeds a YAML resource limit.", error);
        } catch (Exception error) when (error is YamlException or JsonException) {
            throw new FormatException("The Jupytext metadata must contain valid YAML within the notebook limits.", error);
        }
    }

    private static PercentMarker ParseMarker(ReadOnlySpan<char> options, NotebookImportContext context, int lineNumber) {
        var metadata = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        int titleEnd = options.Length;
        if (options.Length > 0) {
            int depth = 0;
            while (depth < options.Length && options[depth] == '%') depth++;
            if (depth > 0) {
                using var number = JsonDocument.Parse(depth.ToString(System.Globalization.CultureInfo.InvariantCulture));
                metadata["cell_depth"] = number.RootElement.Clone();
                options = options.Slice(depth).TrimStart();
            }
            titleEnd = ReadMarkerMetadata(options, metadata, context, lineNumber);
        }

        string title = options.Slice(0, titleEnd).Trim().ToString();
        string? kind = MetadataString(metadata, "cell_type");
        foreach (string label in CellLabels) {
            int tag = title.IndexOf(label, StringComparison.Ordinal);
            if (tag < 0) continue;
            kind = label.Substring(1, label.Length - 2);
            title = title.Remove(tag, label.Length).Trim();
            if (kind == "md") metadata["region_name"] = NotebookImportContext.JsonString(kind);
            break;
        }

        if (title.Length != 0) metadata["title"] = NotebookImportContext.JsonString(title);
        string? id = MetadataString(metadata, "id") ?? (metadata.ContainsKey("id") ? string.Empty : null);
        string? language = MetadataString(metadata, "language");
        if (string.IsNullOrWhiteSpace(language)) language = null;
        NotebookCellKind cellKind = kind switch {
            "markdown" or "md" => NotebookCellKind.Markdown,
            "raw" => NotebookCellKind.Raw,
            null or "code" when IsActiveLanguage(metadata, "ipynb") => NotebookCellKind.Code,
            _ => NotebookCellKind.Raw
        };
        if (kind != null && kind is not ("markdown" or "md" or "raw" or "code")) {
            context.Report("notebook-import-cell-type", "The Jupytext cell type was imported as raw source.", LineLocation(lineNumber),
                OfficeConversionLossKind.Approximation);
        }
        if (kind is "markdown" or "md" or "raw" or "code") metadata.Remove("cell_type");
        return new PercentMarker(cellKind, id, language, metadata);
    }

    private static readonly string[] CellLabels = ["[markdown]", "[raw]", "[md]"];

    private static int ReadMarkerMetadata(ReadOnlySpan<char> options, Dictionary<string, JsonElement> metadata,
        NotebookImportContext context, int line) {
        for (int i = 0; i < options.Length;) {
            context.CancellationToken.ThrowIfCancellationRequested();
            int start = i;
            if (options[i] == '{') {
                try {
                    using var value = JsonDocument.Parse(options.Slice(i).ToString(), JsonOptions(context));
                    foreach (JsonProperty property in value.RootElement.EnumerateObject()) metadata[property.Name] = property.Value.Clone();
                } catch (JsonException) {
                    PreserveMarkerMetadata(options.Slice(i), metadata, context, line);
                }
                return i;
            }
            while (i < options.Length && !char.IsWhiteSpace(options[i]) && options[i] != '=') i++;
            if (i < options.Length && options[i] == '=' && IsMetadataKey(options.Slice(start, i - start))) {
                string key = options.Slice(start, i - start).ToString();
                ReadAssignments(options.Slice(i + 1), key, metadata, context, line);
                return start;
            }
            while (i < options.Length && !char.IsWhiteSpace(options[i])) i++;
            while (i < options.Length && char.IsWhiteSpace(options[i])) i++;
        }

        return options.Length;
    }

    private static void ReadAssignments(ReadOnlySpan<char> assignments, string key, Dictionary<string, JsonElement> metadata,
        NotebookImportContext context, int line) {
        while (true) {
            context.CancellationToken.ThrowIfCancellationRequested();
            assignments = assignments.TrimStart();
            int end = AssignmentValueEnd(assignments);
            if (end == 0) {
                PreserveMarkerMetadata(assignments, metadata, context, line);
                return;
            }
            try {
                using JsonDocument value = ParseMarkerValue(assignments.Slice(0, end).ToString(), context);
                metadata[key] = value.RootElement.Clone();
            } catch (FormatException) {
                PreserveMarkerMetadata(assignments, metadata, context, line);
                return;
            }
            assignments = assignments.Slice(end).TrimStart();
            if (assignments.IsEmpty) return;
            end = assignments.IndexOf('=');
            ReadOnlySpan<char> nextKey = end > 0 ? assignments.Slice(0, end).TrimEnd() : default;
            if (!IsMetadataKey(nextKey)) {
                PreserveMarkerMetadata(assignments, metadata, context, line);
                return;
            }
            key = nextKey.ToString();
            assignments = assignments.Slice(end + 1);
        }
    }

    private static JsonDocument ParseMarkerValue(string text, NotebookImportContext context) {
        string json = text switch { "True" => "true", "False" => "false", "None" => "null", _ => text };
        try {
            return JsonDocument.Parse(json, JsonOptions(context));
        } catch (JsonException error) {
            if (json.Length > 2 && json[0] is '\'' or '{' or '[') return ParsePercentYaml(json, context);
            throw new FormatException("The Jupytext cell metadata value is not supported.", error);
        }
    }

    private static int AssignmentValueEnd(ReadOnlySpan<char> value) {
        int depth = 0;
        char quote = '\0';
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (quote != '\0') {
                if (c == '\\') i++;
                else if (c == quote) quote = '\0';
            } else if (c is '\'' or '"') quote = c;
            else if (c is '[' or '{') depth++;
            else if (c is ']' or '}') depth--;
            else if (depth == 0 && char.IsWhiteSpace(c)) return i;
        }
        return value.Length;
    }

    private static bool IsMetadataKey(ReadOnlySpan<char> value) {
        if (value.IsEmpty || (!char.IsLetter(value[0]) && value[0] != '_')) return false;
        foreach (char c in value) {
            if (!char.IsLetterOrDigit(c) && c is not ('_' or '-' or '.')) return false;
        }
        return true;
    }

    private static void PreserveMarkerMetadata(ReadOnlySpan<char> text, Dictionary<string, JsonElement> metadata,
        NotebookImportContext context, int line) {
        metadata["officeimo_jupytext_marker_metadata"] = NotebookImportContext.JsonString(text.ToString());
        context.Report("notebook-import-script-metadata", "Unparsed Jupytext cell marker metadata was retained as text.",
            LineLocation(line), OfficeConversionLossKind.Approximation);
    }

    private static string? MetadataString(Dictionary<string, JsonElement> metadata, string key) =>
        metadata.TryGetValue(key, out JsonElement value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;

    private static bool IsActiveLanguage(Dictionary<string, JsonElement> metadata, string language) {
        string script = language switch { "python" => "py", "julia" => "jl", _ => language };
        if (metadata.TryGetValue("run_control", out JsonElement runControl)
            && ReadObjectProperty(runControl, "frozen").ValueKind == JsonValueKind.True) return language == "ipynb";
        if (metadata.TryGetValue("tags", out JsonElement tags) && tags.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement tag in tags.EnumerateArray()) {
                string? value = tag.ValueKind == JsonValueKind.String ? tag.GetString() : null;
                if (value != null && value.StartsWith("active-", StringComparison.Ordinal)) {
                    return Array.Exists(value.Substring(7).Split('-'), part => string.Equals(part, script, StringComparison.OrdinalIgnoreCase));
                }
            }
        }
        string? active = MetadataString(metadata, "active");
        return active == null || Array.Exists(active.Split('.', ','), part => string.Equals(part, script, StringComparison.OrdinalIgnoreCase));
    }
}
