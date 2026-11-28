using System.Globalization;
using System.Text.Json;
using OfficeIMO.Yaml;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookDeepnoteImporter {
    private static void ReadOutputs(YamlSequence outputs, NotebookCell cell, string location, NotebookImportContext context) {
        for (int i = 0; i < outputs.Count; i++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            string outputLocation = location + ".outputs[" + i.ToString(CultureInfo.InvariantCulture) + "]";
            NotebookOutput? output = outputs[i] is YamlMapping saved ? ReadOutput(saved, outputLocation, context) : null;
            if (output != null) {
                context.AddOutput(cell, output);
            } else {
                context.Report("notebook-deepnote-output", "The saved output is malformed or uses an unsupported output type.", outputLocation);
            }
        }
    }

    private static NotebookOutput? ReadOutput(YamlMapping output, string location, NotebookImportContext context) {
        switch (Text(output, "output_type")) {
            case "stream":
                string? name = Text(output, "name");
                return name is "stdout" or "stderr" && TryReadStrings(Get(output, "text"), out IReadOnlyList<string> text)
                    ? NotebookOutput.CreateStream(name, string.Concat(text))
                    : null;
            case "error":
                string? errorName = Text(output, "ename");
                string? errorValue = Text(output, "evalue");
                return errorName != null && errorValue != null && TryReadStrings(Get(output, "traceback"), out IReadOnlyList<string> traceback)
                    ? NotebookOutput.CreateError(errorName, errorValue, traceback)
                    : null;
            case "display_data":
            case "execute_result": {
                if (Get(output, "data") is not YamlMapping data) return null;
                using JsonDocument dataJson = ParseMap(data, context);
                Dictionary<string, JsonElement> bundle = NotebookJson.ReadMap(dataJson.RootElement);
                YamlNode? metadataNode = Get(output, "metadata");
                using JsonDocument? metadataJson = metadataNode is YamlMapping map ? ParseMap(map, context) : null;
                var metadata = metadataJson != null ? NotebookJson.ReadMap(metadataJson.RootElement) : null;
                if (metadataNode is not (null or YamlMapping)) {
                    context.Report("notebook-deepnote-output-metadata", "The saved output metadata is not a mapping.", location + ".metadata");
                }
                return Text(output, "output_type") == "display_data"
                    ? NotebookOutput.CreateDisplayData(bundle, metadata)
                    : NotebookOutput.CreateExecuteResult(bundle, Count(output, "execution_count", location, context), metadata);
            }
            default:
                return null;
        }
    }

    private static JsonDocument ParseMap(YamlMapping mapping, NotebookImportContext context) {
        // Reuse the YAML driver's schema-aware JSON conversion for nested MIME values, including
        // producer-specific JSON. Keep all alternatives so later exporters can select a safe view.
        return JsonDocument.Parse(mapping.ToJsonData().ToString(),
            new JsonDocumentOptions { MaxDepth = context.Limits.MaxJsonDepth });
    }

    private static bool TryReadStrings(YamlNode? node, out IReadOnlyList<string> result) {
        if (node is YamlScalar { IsNull: false } scalar) {
            result = new[] { scalar.AsString() };
            return true;
        }
        if (node is YamlSequence sequence) {
            var values = new List<string>(sequence.Count);
            foreach (YamlNode item in sequence) {
                if (item is not YamlScalar { IsNull: false } fragment) {
                    result = Array.Empty<string>();
                    return false;
                }
                values.Add(fragment.AsString());
            }
            result = values;
            return true;
        }
        result = Array.Empty<string>();
        return false;
    }

    private static int? Count(YamlMapping source, string name, string location, NotebookImportContext context) {
        YamlNode? count = Get(source, name);
        if (count is null or YamlScalar { IsNull: true }) return null;
        if (count is YamlScalar scalar && scalar.TryGetValue(out int value) && value >= 0) return value;
        context.Report("notebook-deepnote-execution-count", "The saved execution count is not a nonnegative integer.", location + "." + name);
        return null;
    }
}
