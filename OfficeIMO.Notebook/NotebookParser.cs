namespace OfficeIMO.Notebook;

internal static class NotebookParser {
    internal static List<JsonElement> ReadSourceCells(JsonElement root, int major, int maxCells, bool limitProjection,
        CancellationToken cancellationToken, List<OfficeDiagnostic> diagnostics, out int sourceCellCount) {
        var cells = new List<JsonElement>();
        sourceCellCount = 0;
        if (major == 4) {
            sourceCellCount = AppendSourceCells(NotebookJson.Property(root, "cells"), cells, maxCells, limitProjection, cancellationToken);
            return cells;
        }

        JsonElement worksheets = NotebookJson.Property(root, "worksheets");
        if (worksheets.ValueKind != JsonValueKind.Array) throw new FormatException("A version 3 notebook must contain a worksheets array.");
        foreach (JsonElement worksheet in worksheets.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            sourceCellCount += AppendSourceCells(NotebookJson.Property(worksheet, "cells"), cells, maxCells, limitProjection, cancellationToken);
        }
        diagnostics.Add(new OfficeDiagnostic("notebook-v3-normalized", OfficeDiagnosticSeverity.Warning,
            "Notebook v3 worksheets were flattened into a v4-style cell sequence for export.", format: "Jupyter Notebook",
            lossKind: OfficeConversionLossKind.Approximation));
        return cells;
    }

    internal static IEnumerable<NotebookCell> ReadCells(List<JsonElement> sourceCells, int major, int minor,
        NotebookLoadOptions options, int? projectionOutputLimit, CancellationToken cancellationToken, List<OfficeDiagnostic> diagnostics) {
        HashSet<string> reserved = FindReservedIds(sourceCells, cancellationToken);
        var used = new HashSet<string>(StringComparer.Ordinal);
        int totalOutputs = 0;
        for (int index = 0; index < sourceCells.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            JsonElement source = sourceCells[index];
            if (source.ValueKind != JsonValueKind.Object) {
                diagnostics.Add(new OfficeDiagnostic("notebook-invalid-cell", OfficeDiagnosticSeverity.Warning,
                    "An entry in the notebook cell array is not an object and was omitted.", $"cell[{index}]",
                    "Jupyter Notebook", OfficeConversionLossKind.Omission));
                continue;
            }

            string id = ReadId(source, index, reserved, used, major, minor, diagnostics);
            string name = NotebookJson.String(source, "cell_type") ?? string.Empty;
            bool legacy = major == 3;
            NotebookCellKind kind = name switch {
                "markdown" => NotebookCellKind.Markdown,
                "code" => NotebookCellKind.Code,
                "raw" => NotebookCellKind.Raw,
                "heading" or "html" when legacy => NotebookCellKind.Markdown,
                _ => NotebookCellKind.Unknown
            };
            JsonElement cellSource = NotebookJson.Property(source, legacy && kind == NotebookCellKind.Code ? "input" : "source");
            string? normalized = legacy && name == "heading" ? ReadLegacyHeading(source, cellSource) : null;
            var metadata = NotebookJson.ReadMap(NotebookJson.Property(source, "metadata"));
            if (legacy && kind == NotebookCellKind.Code) MergeLegacyMetadata(source, metadata);
            string? language = ReadLanguage(source, metadata);
            int? count = NotebookJson.Integer(source, legacy ? "prompt_number" : "execution_count");
            int? outputLimit = projectionOutputLimit.HasValue && kind != NotebookCellKind.Code ? 0 : projectionOutputLimit;
            var outputs = NotebookOutputParser.Read(source, legacy, options, outputLimit, ref totalOutputs, cancellationToken);
            yield return new NotebookCell(kind, name, id, index, source, cellSource, normalized, language, count, metadata, outputs);
        }
    }

    private static int AppendSourceCells(JsonElement array, List<JsonElement> cells, int maxCells,
        bool limitProjection, CancellationToken cancellationToken) {
        if (array.ValueKind != JsonValueKind.Array) throw new FormatException("Notebook JSON must contain a cells array.");
        int count = array.GetArrayLength();
        if (!limitProjection && count > maxCells - cells.Count) throw Limit(nameof(NotebookLoadOptions.MaxCells), maxCells);
        var items = array.EnumerateArray();
        while (cells.Count < maxCells && items.MoveNext()) {
            cancellationToken.ThrowIfCancellationRequested();
            cells.Add(items.Current);
        }
        return count;
    }

    private static HashSet<string> FindReservedIds(List<JsonElement> cells, CancellationToken cancellationToken) {
        var reserved = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement cell in cells) {
            cancellationToken.ThrowIfCancellationRequested();
            string? id = NotebookJson.String(cell, "id");
            if (IsValidId(id)) reserved.Add(id!);
        }
        return reserved;
    }

    private static string ReadId(JsonElement cell, int index, HashSet<string> reserved, HashSet<string> used,
        int major, int minor, List<OfficeDiagnostic> diagnostics) {
        string? original = NotebookJson.String(cell, "id");
        if (IsValidId(original) && used.Add(original!)) return original!;
        string prefix = "cell-" + (index + 1).ToString(CultureInfo.InvariantCulture);
        string id = prefix;
        for (int suffix = 2; reserved.Contains(id) || !used.Add(id); suffix++) {
            id = prefix + "-" + suffix.ToString(CultureInfo.InvariantCulture);
        }

        if (original != null || (major == 4 && minor >= 5)) {
            diagnostics.Add(new OfficeDiagnostic("notebook-cell-id-repaired", OfficeDiagnosticSeverity.Warning,
                "A missing, invalid, or duplicate notebook cell ID was replaced.", id, "Jupyter Notebook"));
        }
        return id;
    }

    internal static bool IsValidId(string? id) =>
        id is { Length: >= 1 and <= 64 } && id.All(static character =>
            character is >= '0' and <= '9' or >= 'A' and <= 'Z' or >= 'a' and <= 'z' or '-' or '_');

    private static string? ReadLanguage(JsonElement source, Dictionary<string, JsonElement> metadata) {
        JsonElement originalMetadata = NotebookJson.Property(source, "metadata");
        string? language = NotebookJson.String(source, "language") ??
            (metadata.TryGetValue("language", out JsonElement languageValue) && languageValue.ValueKind == JsonValueKind.String
                ? languageValue.GetString() : null) ??
            NotebookJson.String(NotebookJson.Property(originalMetadata, "polyglot_notebook"), "kernelName") ??
            NotebookJson.String(NotebookJson.Property(originalMetadata, "dotnet_interactive"), "language");
        if (!string.IsNullOrWhiteSpace(language)) return language;
        string? editorLanguage = NotebookJson.String(NotebookJson.Property(originalMetadata, "vscode"), "languageId");
        return editorLanguage == "polyglot-notebook" ? null : editorLanguage;
    }

    private static void MergeLegacyMetadata(JsonElement source, Dictionary<string, JsonElement> metadata) {
        if (NotebookJson.Property(source, "collapsed") is { ValueKind: JsonValueKind.True or JsonValueKind.False } collapsed) {
            metadata["collapsed"] = collapsed;
        }
        if (!metadata.ContainsKey("language") &&
            NotebookJson.Property(source, "language") is { ValueKind: JsonValueKind.String } language) {
            metadata["language"] = language;
        }
    }

    private static string ReadLegacyHeading(JsonElement source, JsonElement text) {
        int level = Math.Min(Math.Max(NotebookJson.Integer(source, "level") ?? 1, 1), 6);
        string content = NotebookJson.ReadText(text).Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
        return new string('#', level) + " " + content;
    }

    internal static InvalidDataException Limit(string name, int value) =>
        new($"Notebook input exceeds the configured {name} limit of {value.ToString(CultureInfo.InvariantCulture)}.");
}
