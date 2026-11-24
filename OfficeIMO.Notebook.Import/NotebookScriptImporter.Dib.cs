using System.Text;
using System.Text.Json;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    internal static void ReadDib(ReadOnlyMemory<byte> input, NotebookImportContext context) {
        var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["csharp"] = "csharp", ["cs"] = "csharp", ["c#"] = "csharp",
            ["fsharp"] = "fsharp", ["fs"] = "fsharp", ["f#"] = "fsharp",
            ["pwsh"] = "pwsh", ["powershell"] = "pwsh",
            ["javascript"] = "javascript", ["js"] = "javascript",
            ["markdown"] = "markdown", ["md"] = "markdown",
            ["html"] = "html", ["mermaid"] = "mermaid", ["http"] = "http",
            ["sql"] = "sql", ["kql"] = "kql", ["value"] = "value"
        };
        using var lines = OpenScript(input);
        string? line = lines.ReadLine();
        while (line is { Length: 0 }) {
            context.CancellationToken.ThrowIfCancellationRequested();
            line = lines.ReadLine();
        }

        string defaultKernel = "csharp";
        string kernel = defaultKernel;
        bool readMetadata = false;
        bool selectedKernel = false;
        var source = new StringBuilder();
        int trailingEmpty = 0;
        while (line != null) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (!readMetadata && line == "#!meta") {
                var header = new StringBuilder();
                while ((line = lines.ReadLine()) != null && !line.StartsWith("#!", StringComparison.Ordinal)) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    header.AppendLine(line);
                }
                defaultKernel = ReadDibMetadata(header.ToString(), aliases, context) ?? defaultKernel;
                if (!selectedKernel && source.Length == 0) kernel = defaultKernel;
                readMetadata = true;
                if (line == null) break;
            }

            if (line.StartsWith("#!", StringComparison.Ordinal) && aliases.TryGetValue(line.Substring(2), out string? selected)) {
                AddDibCell(context, kernel, source, trailingEmpty);
                source.Clear();
                trailingEmpty = 0;
                kernel = selected;
                selectedKernel = true;
            } else if (source.Length != 0 || line.Length != 0) {
                if (source.Length != 0) source.Append('\n');
                source.Append(line);
                trailingEmpty = line.Length == 0 ? trailingEmpty + 1 : 0;
            }
            line = lines.ReadLine();
        }

        SetLanguageInfo(context.Document, defaultKernel);
        AddDibCell(context, kernel, source, trailingEmpty);
        if (context.Document.Cells.Count == 0) {
            bool markdown = kernel == "markdown";
            context.AddCell(markdown ? NotebookCellKind.Markdown : NotebookCellKind.Code, string.Empty, markdown ? null : kernel);
        }
    }

    private static void AddDibCell(NotebookImportContext context, string kernel, StringBuilder source, int trailingEmpty) {
        int length = source.Length - trailingEmpty;
        if (length == 0) return;
        bool markdown = kernel == "markdown";
        context.AddCell(markdown ? NotebookCellKind.Markdown : NotebookCellKind.Code, source.ToString(0, length), markdown ? null : kernel);
    }

    private static string? ReadDibMetadata(string header, IDictionary<string, string> aliases, NotebookImportContext context) {
        JsonDocument metadata;
        try {
            metadata = JsonDocument.Parse(header, JsonOptions(context));
        } catch (JsonException error) {
            throw new FormatException("The Polyglot notebook's #!meta block must contain a valid JSON object within the notebook limits.", error);
        }
        using (metadata) {
            JsonElement root = metadata.RootElement;
            if (root.ValueKind != JsonValueKind.Object) throw new FormatException("The Polyglot notebook's #!meta block must be a JSON object.");
            foreach (JsonProperty property in root.EnumerateObject()) {
                context.CancellationToken.ThrowIfCancellationRequested();
                context.Document.SetMetadata(property.Name, property.Value);
            }

            JsonElement kernelInfo = ReadObjectProperty(root, "kernelInfo");
            if (kernelInfo.ValueKind == JsonValueKind.Undefined) {
                kernelInfo = ReadObjectProperty(ReadObjectProperty(root, "polyglot_notebook"), "kernelInfo");
            }
            JsonElement items = ReadObjectProperty(kernelInfo, "items");
            if (items.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement item in items.EnumerateArray()) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    string? name = ReadStringProperty(item, "name");
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    string kernel = aliases.TryGetValue(name!, out string? known) ? known : name!;
                    aliases[name!] = kernel;
                    JsonElement kernelAliases = ReadObjectProperty(item, "aliases");
                    if (kernelAliases.ValueKind != JsonValueKind.Array) continue;
                    foreach (JsonElement alias in kernelAliases.EnumerateArray()) {
                        context.CancellationToken.ThrowIfCancellationRequested();
                        if (alias.ValueKind == JsonValueKind.String && alias.GetString() is { Length: > 0 } value &&
                            !aliases.ContainsKey(value)) aliases[value] = kernel;
                    }
                }
            }

            string? defaultName = ReadStringProperty(kernelInfo, "defaultKernelName");
            return defaultName != null && aliases.TryGetValue(defaultName, out string? selected) ? selected : defaultName;
        }
    }
}
