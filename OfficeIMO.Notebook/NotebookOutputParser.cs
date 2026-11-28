namespace OfficeIMO.Notebook;

internal static class NotebookOutputParser {
    internal static List<NotebookOutput> Read(JsonElement source, bool legacy, NotebookLoadOptions options,
        int? projectionLimit, ref int totalOutputs, CancellationToken cancellationToken) {
        var outputs = new List<NotebookOutput>();
        JsonElement array = NotebookJson.Property(source, "outputs");
        if (array.ValueKind != JsonValueKind.Array) return outputs;
        int count = array.GetArrayLength();
        if (projectionLimit.HasValue) {
            count = Math.Min(count, projectionLimit.Value);
        } else {
            if (count > options.MaxOutputsPerCell) throw NotebookParser.Limit(nameof(options.MaxOutputsPerCell), options.MaxOutputsPerCell);
            if (count > options.MaxTotalOutputs - totalOutputs) throw NotebookParser.Limit(nameof(options.MaxTotalOutputs), options.MaxTotalOutputs);
            totalOutputs += count;
        }

        var items = array.EnumerateArray();
        while (outputs.Count < count && items.MoveNext()) {
            cancellationToken.ThrowIfCancellationRequested();
            outputs.Add(Parse(items.Current, legacy, options.MaxJsonDepth));
        }
        return outputs;
    }

    private static NotebookOutput Parse(JsonElement output, bool legacy, int maxJsonDepth) {
        string name = NotebookJson.String(output, "output_type") ?? string.Empty;
        NotebookOutputKind kind = name switch {
            "stream" => NotebookOutputKind.Stream,
            "display_data" => NotebookOutputKind.DisplayData,
            "execute_result" => NotebookOutputKind.ExecuteResult,
            "error" => NotebookOutputKind.Error,
            "pyout" when legacy => NotebookOutputKind.ExecuteResult,
            "pyerr" when legacy => NotebookOutputKind.Error,
            _ => NotebookOutputKind.Unknown
        };
        bool legacyDisplay = legacy && kind is NotebookOutputKind.DisplayData or NotebookOutputKind.ExecuteResult;
        Dictionary<string, JsonElement> data = legacyDisplay ? ReadLegacyData(output, maxJsonDepth) :
            NotebookJson.ReadMap(NotebookJson.Property(output, "data"));
        var metadata = NotebookJson.ReadMap(NotebookJson.Property(output, "metadata"), legacyDisplay ? MimeName : null);
        return new NotebookOutput(kind, name, output,
            NotebookJson.String(output, legacy ? "stream" : "name") ?? (legacy && kind == NotebookOutputKind.Stream ? "stdout" : null),
            NotebookJson.Property(output, "text"), null, NotebookJson.String(output, "ename"),
            NotebookJson.String(output, "evalue"), NotebookJson.Property(output, "traceback"), null,
            NotebookJson.Integer(output, legacy ? "prompt_number" : "execution_count"), data, metadata, legacy);
    }

    private static Dictionary<string, JsonElement> ReadLegacyData(JsonElement output, int maxJsonDepth) {
        var data = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (JsonProperty property in output.EnumerateObject()) {
            if (property.Name is "output_type" or "prompt_number" or "metadata") continue;
            string name = MimeName(property.Name);
            if (name == "application/json" && property.Value.ValueKind == JsonValueKind.String) {
                try {
                    using JsonDocument value = JsonDocument.Parse(property.Value.GetString()!, new JsonDocumentOptions { MaxDepth = maxJsonDepth });
                    data[name] = value.RootElement.Clone();
                } catch (JsonException) {
                    data[name] = property.Value;
                }
            } else {
                data[name] = property.Value;
            }
        }
        return data;
    }

    private static string MimeName(string alias) => alias switch {
        "text" => "text/plain",
        "html" => "text/html",
        "svg" => "image/svg+xml",
        "png" => "image/png",
        "jpeg" => "image/jpeg",
        "latex" => "text/latex",
        "json" => "application/json",
        "javascript" => "application/javascript",
        _ => alias
    };
}
