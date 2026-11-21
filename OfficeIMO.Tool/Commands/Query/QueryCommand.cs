using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeIMO.Toml;
using OfficeIMO.Yaml;

namespace OfficeIMO.Tool.Commands.Query;

/// <summary>Runs one path query against a YAML, JSON or TOML file through the engines' <c>Query</c> methods.</summary>
internal static class QueryCommand {
    internal const string Usage = """
OfficeIMO.Tool - query YAML, JSON and TOML

Usage:
  officeimo query <file> <query> [--format text|json] [--input-format yaml|json|toml]

The same query reads all three formats:
  services[*].name                      every item's name
  ..image                               every image key at any depth
  services[?(@.port >= 8000)].name      items whose port is at least 8000
  package[?(@.name == 'serde')].version an array-of-tables entry by field
  $[?(@.kind == 'Service')].metadata    documents of a multi-document YAML stream

Options:
  --format <text|json>         text (default): one line per match, path, line:column and value.
                               json: an array of {path, line, column, kind, value}.
  --input-format <yaml|json|toml>  Read the file as this format instead of by extension.

Paths in the output select the same node with YamlDocument.Select or TomlDocument.Select.
Collections print as compact JSON. Positions are one-based; 0 means the node has no source position.
""";

    internal static IEnumerable<string> SupportedExtensions() => ["yaml", "yml", "json", "toml"];

    private enum InputFormat {
        Yaml,
        Toml
    }

    private sealed record Match(string Path, int Line, int Column, string Kind, Action<Utf8JsonWriter> WriteValue, string? Text);

    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken cancellationToken = default) {
        string? file = null, query = null, format = "text", inputFormat = null;
        for (int i = 0; i < args.Length; i++) {
            string arg = args[i];
            if (arg is "--help" or "-h" or "help") {
                await output.WriteLineAsync(Usage).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            if (arg is "--format" or "--input-format") {
                if (i + 1 >= args.Length) return await UsageErrorAsync(error, arg + " needs a value.").ConfigureAwait(false);
                if (arg == "--format") format = args[++i];
                else inputFormat = args[++i];
            } else if (arg.StartsWith("--", StringComparison.Ordinal)) {
                return await UsageErrorAsync(error, "Unknown option: " + arg).ConfigureAwait(false);
            } else if (file is null) {
                file = arg;
            } else if (query is null) {
                query = arg;
            } else {
                return await UsageErrorAsync(error, "Unexpected argument: " + arg).ConfigureAwait(false);
            }
        }
        if (file is null || query is null) return await UsageErrorAsync(error, "A file and a query are required.").ConfigureAwait(false);
        if (format is not ("text" or "json")) return await UsageErrorAsync(error, "--format must be text or json.").ConfigureAwait(false);
        InputFormat kind;
        string detected = (inputFormat ?? Path.GetExtension(file).TrimStart('.')).ToLowerInvariant();
        switch (detected) {
            case "yaml": case "yml": case "json": kind = InputFormat.Yaml; break;
            case "toml": kind = InputFormat.Toml; break;
            default:
                return await UsageErrorAsync(error, inputFormat is null
                    ? "Cannot tell the format of '" + file + "'; use --input-format yaml, json or toml."
                    : "--input-format must be yaml, json or toml.").ConfigureAwait(false);
        }
        if (!File.Exists(file)) {
            await error.WriteLineAsync("Input was not found: " + file).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        }

        List<Match> matches;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            matches = kind == InputFormat.Yaml
                ? YamlDocument.Load(file, cancellationToken: cancellationToken).Query(query).Select(FromYaml).ToList()
                : TomlDocument.Load(file).Query(query).Select(FromToml).ToList();
        } catch (OperationCanceledException) {
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (ArgumentException exception) {
            // A malformed query: the message names the position; the parameter suffix means nothing on a command line.
            string message = exception.ParamName is null ? exception.Message : exception.Message.Replace(" (Parameter '" + exception.ParamName + "')", string.Empty);
            await error.WriteLineAsync(message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (Exception exception) when (exception is YamlException or TomlException or IOException or UnauthorizedAccessException) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        }

        if (format == "json") {
            await output.WriteLineAsync(ToJson(matches)).ConfigureAwait(false);
        } else {
            foreach (Match match in matches) {
                await output.WriteLineAsync(match.Path + "\t" + match.Line.ToString(CultureInfo.InvariantCulture) + ":" +
                    match.Column.ToString(CultureInfo.InvariantCulture) + "\t" + (match.Text is null ? Compact(match.WriteValue) : Escape(match.Text))).ConfigureAwait(false);
            }
        }
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task<int> UsageErrorAsync(TextWriter error, string message) {
        await error.WriteLineAsync(message).ConfigureAwait(false);
        await error.WriteLineAsync(Usage).ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Usage;
    }

    private static string ToJson(List<Match> matches) {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer, new JsonWriterOptions { Indented = true })) {
            writer.WriteStartArray();
            foreach (Match match in matches) {
                writer.WriteStartObject();
                writer.WriteString("path", match.Path);
                writer.WriteNumber("line", match.Line);
                writer.WriteNumber("column", match.Column);
                writer.WriteString("kind", match.Kind);
                writer.WritePropertyName("value");
                match.WriteValue(writer);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    private static string Compact(Action<Utf8JsonWriter> write) {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer)) write(writer);
        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    /// <summary>Keeps one match per line: backslashes, tabs and line breaks are escaped.</summary>
    private static string Escape(string text) => text.Replace("\\", "\\\\").Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n");

    private static Match FromYaml(YamlQueryMatch match) {
        YamlNode node = match.Node;
        YamlNode value = node is YamlAlias alias ? alias.Target : node;
        string kind = value switch {
            YamlMapping => "mapping",
            YamlSequence => "sequence",
            YamlScalar scalar => ScalarKind(scalar),
            _ => "scalar"
        };
        string? text = value is YamlScalar s && kind != "null" ? s.AsString() : null;
        return new Match(match.Path, node.Start.Line, node.Start.Column, kind, writer => WriteYaml(writer, value, 0), text);
    }

    private static string ScalarKind(YamlScalar scalar) => scalar.Kind switch {
        YamlScalarKind.Null => "null",
        YamlScalarKind.Boolean => "boolean",
        YamlScalarKind.Integer or YamlScalarKind.Float => "number",
        _ => "string"
    };

    private static void WriteYaml(Utf8JsonWriter writer, YamlNode node, int depth) {
        if (node is YamlAlias alias) node = alias.Target;
        if (depth > 64) throw new InvalidOperationException("The value nests more than 64 levels deep through aliases.");
        switch (node) {
            case YamlMapping mapping:
                writer.WriteStartObject();
                foreach (YamlMappingEntry entry in mapping) {
                    YamlNode key = entry.Key is YamlAlias keyAlias ? keyAlias.Target : entry.Key;
                    writer.WritePropertyName(key is YamlScalar scalarKey ? scalarKey.AsString() : key.ToString());
                    WriteYaml(writer, entry.Value, depth + 1);
                }
                writer.WriteEndObject();
                break;
            case YamlSequence sequence:
                writer.WriteStartArray();
                foreach (YamlNode item in sequence) WriteYaml(writer, item, depth + 1);
                writer.WriteEndArray();
                break;
            case YamlScalar scalar:
                switch (scalar.Kind) {
                    case YamlScalarKind.Null: writer.WriteNullValue(); break;
                    case YamlScalarKind.Boolean when scalar.TryGetValue(out bool boolean): writer.WriteBooleanValue(boolean); break;
                    case YamlScalarKind.Integer when scalar.TryGetValue(out long integer): writer.WriteNumberValue(integer); break;
                    case YamlScalarKind.Integer or YamlScalarKind.Float when scalar.TryGetValue(out double number) && IsFinite(number): writer.WriteNumberValue(number); break;
                    default: writer.WriteStringValue(scalar.AsString()); break;
                }
                break;
        }
    }

    private static Match FromToml(TomlQueryMatch match) {
        TomlNode node = match.Node;
        string kind = node.Kind switch {
            TomlNodeKind.Table => "mapping",
            TomlNodeKind.Array or TomlNodeKind.TableArray => "sequence",
            TomlNodeKind.Integer or TomlNodeKind.Float => "number",
            TomlNodeKind.Boolean => "boolean",
            _ => "string"
        };
        string? text = node is TomlValue value ? System.Convert.ToString(value.Value, CultureInfo.InvariantCulture) : null;
        if (node is TomlValue { Value: bool flag }) text = flag ? "true" : "false";
        return new Match(match.Path, node.Start.Line, node.Start.Column, kind, writer => WriteToml(writer, node), text);
    }

    private static void WriteToml(Utf8JsonWriter writer, TomlNode node) {
        switch (node) {
            case TomlTable table:
                writer.WriteStartObject();
                foreach (KeyValuePair<string, TomlNode> entry in table) {
                    writer.WritePropertyName(entry.Key);
                    WriteToml(writer, entry.Value);
                }
                writer.WriteEndObject();
                break;
            case TomlArray array:
                writer.WriteStartArray();
                foreach (TomlNode item in array) WriteToml(writer, item);
                writer.WriteEndArray();
                break;
            case TomlTableArray tables:
                writer.WriteStartArray();
                foreach (TomlTable item in tables) WriteToml(writer, item);
                writer.WriteEndArray();
                break;
            case TomlValue value:
                switch (value.Value) {
                    case string text: writer.WriteStringValue(text); break;
                    case long integer: writer.WriteNumberValue(integer); break;
                    case double number when IsFinite(number): writer.WriteNumberValue(number); break;
                    case bool boolean: writer.WriteBooleanValue(boolean); break;
                    default: writer.WriteStringValue(System.Convert.ToString(value.Value, CultureInfo.InvariantCulture)); break;
                }
                break;
        }
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
