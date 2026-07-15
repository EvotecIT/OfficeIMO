namespace OfficeIMO.Reader.Notebook;

internal static class NotebookReaderAdapter {
    internal static OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions readerOptions,
        ReaderNotebookOptions notebookOptions,
        CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(path, readerOptions, cancellationToken);
        return BuildDocument(input, readerOptions, notebookOptions, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(
        Stream stream,
        string? sourceName,
        ReaderOptions readerOptions,
        ReaderNotebookOptions notebookOptions,
        CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(
            stream,
            string.IsNullOrWhiteSpace(sourceName) ? "notebook.ipynb" : sourceName,
            readerOptions,
            cancellationToken);
        return BuildDocument(input, readerOptions, notebookOptions, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocument(
        ReaderAdapterInputSnapshot input,
        ReaderOptions readerOptions,
        ReaderNotebookOptions notebookOptions,
        CancellationToken cancellationToken) {
        using JsonDocument json = JsonDocument.Parse(input.Bytes, new JsonDocumentOptions { MaxDepth = 128 });
        JsonElement root = json.RootElement;
        if (root.ValueKind != JsonValueKind.Object ||
            !root.TryGetProperty("cells", out JsonElement cells) ||
            cells.ValueKind != JsonValueKind.Array) {
            throw new FormatException("Jupyter Notebook JSON must contain a cells array.");
        }

        string language = FindLanguage(root);
        var chunks = new List<ReaderChunk>(Math.Min(cells.GetArrayLength(), notebookOptions.MaxCells));
        int inspectedCells = 0;
        int includedCells = 0;
        bool cellLimitReached = false;
        foreach (JsonElement cell in cells.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (inspectedCells >= notebookOptions.MaxCells) {
                cellLimitReached = true;
                break;
            }

            int cellIndex = inspectedCells++;
            if (cell.ValueKind != JsonValueKind.Object) continue;
            string cellType = GetString(cell, "cell_type") ?? "unknown";
            if (string.Equals(cellType, "code", StringComparison.Ordinal) && !notebookOptions.IncludeCodeCells) continue;
            if (cellType is not ("markdown" or "raw" or "code")) continue;

            var warnings = new List<string>();
            string source = GetMultilineText(cell, "source");
            source = Truncate(source, notebookOptions.MaxCellCharacters, "Notebook cell source was truncated.", warnings);
            IReadOnlyList<NotebookOutput> outputs = cellType == "code" && notebookOptions.IncludeOutputs
                ? ReadOutputs(cell, notebookOptions, warnings)
                : Array.Empty<NotebookOutput>();
            string markdown = BuildMarkdown(cellType, cellIndex, language, source, outputs);
            string text = BuildPlainText(source, outputs);
            if (string.IsNullOrWhiteSpace(markdown) && string.IsNullOrWhiteSpace(text)) continue;

            string anchor = "notebook-cell-" + cellIndex.ToString("D4", CultureInfo.InvariantCulture);
            var location = new ReaderLocation {
                Path = input.Source.Path,
                BlockIndex = includedCells,
                SourceBlockIndex = cellIndex,
                SourceBlockKind = "notebook-" + cellType + "-cell",
                BlockAnchor = anchor
            };
            var chunk = new ReaderChunk {
                Id = anchor,
                Kind = ReaderInputKind.Json,
                Location = location,
                Text = text,
                Markdown = markdown,
                Warnings = warnings.Count == 0 ? null : warnings.ToArray()
            };
            DocumentReaderEngine.ApplyAdapterSource(chunk, input, readerOptions.ComputeHashes);
            chunks.Add(chunk);
            includedCells++;
        }

        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Json,
            input.Source,
            new[] { OfficeDocumentReaderBuilderNotebookExtensions.HandlerId, "jupyter.nbformat" });
        result.Source.Title = Path.GetFileName(input.Source.Path ?? "notebook.ipynb");
        result.Metadata = result.Metadata.Concat(BuildMetadata(root, language, cells.GetArrayLength(), includedCells)).ToArray();
        if (cellLimitReached) {
            result.Diagnostics = result.Diagnostics.Concat(new[] {
                new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Category = OfficeDocumentDiagnosticCategory.Limit,
                    Code = "notebook-cell-limit",
                    Message = "Notebook cell inspection stopped at MaxCells.",
                    Source = OfficeDocumentReaderBuilderNotebookExtensions.HandlerId,
                    IsRecoverable = true,
                    Attributes = new Dictionary<string, string> {
                        ["maxCells"] = notebookOptions.MaxCells.ToString(CultureInfo.InvariantCulture),
                        ["totalCells"] = cells.GetArrayLength().ToString(CultureInfo.InvariantCulture)
                    }
                }
            }).ToArray();
        }
        return result;
    }

    private static IReadOnlyList<NotebookOutput> ReadOutputs(
        JsonElement cell,
        ReaderNotebookOptions options,
        List<string> warnings) {
        if (!cell.TryGetProperty("outputs", out JsonElement outputs) || outputs.ValueKind != JsonValueKind.Array) {
            return Array.Empty<NotebookOutput>();
        }

        var projected = new List<NotebookOutput>(Math.Min(outputs.GetArrayLength(), options.MaxOutputsPerCell));
        int totalCharacters = 0;
        int inspected = 0;
        foreach (JsonElement output in outputs.EnumerateArray()) {
            if (inspected++ >= options.MaxOutputsPerCell) {
                warnings.Add("Notebook outputs were truncated at MaxOutputsPerCell.");
                break;
            }
            NotebookOutput? value = ProjectOutput(output);
            if (!value.HasValue || string.IsNullOrWhiteSpace(value.Value.Text)) continue;
            int remaining = options.MaxOutputCharactersPerCell - totalCharacters;
            if (remaining <= 0) {
                warnings.Add("Notebook output text was truncated at MaxOutputCharactersPerCell.");
                break;
            }
            string text = value.Value.Text;
            if (text.Length > remaining) {
                text = text.Substring(0, remaining);
                warnings.Add("Notebook output text was truncated at MaxOutputCharactersPerCell.");
            }
            projected.Add(new NotebookOutput(text, value.Value.IsMarkdown));
            totalCharacters += text.Length;
            if (text.Length == remaining) break;
        }
        return projected;
    }

    private static NotebookOutput? ProjectOutput(JsonElement output) {
        if (output.ValueKind != JsonValueKind.Object) return null;
        string outputType = GetString(output, "output_type") ?? string.Empty;
        if (string.Equals(outputType, "stream", StringComparison.Ordinal)) {
            return new NotebookOutput(GetMultilineText(output, "text"), false);
        }
        if (string.Equals(outputType, "error", StringComparison.Ordinal)) {
            string traceback = GetMultilineText(output, "traceback");
            if (!string.IsNullOrWhiteSpace(traceback)) return new NotebookOutput(traceback, false);
            string error = string.Join(": ", new[] { GetString(output, "ename"), GetString(output, "evalue") }
                .Where(static value => !string.IsNullOrWhiteSpace(value)));
            return new NotebookOutput(error, false);
        }
        if (output.TryGetProperty("data", out JsonElement data) && data.ValueKind == JsonValueKind.Object) {
            if (data.TryGetProperty("text/markdown", out JsonElement markdown)) {
                return new NotebookOutput(GetMultilineText(markdown), true);
            }
            if (data.TryGetProperty("text/plain", out JsonElement plain)) {
                return new NotebookOutput(GetMultilineText(plain), false);
            }
        }
        return null;
    }

    private static string BuildMarkdown(
        string cellType,
        int cellIndex,
        string language,
        string source,
        IReadOnlyList<NotebookOutput> outputs) {
        if (cellType == "markdown") return source.Trim();
        var markdown = new StringBuilder();
        markdown.Append("### ")
            .Append(cellType == "code" ? "Code" : "Raw")
            .Append(" cell ")
            .AppendLine((cellIndex + 1).ToString(CultureInfo.InvariantCulture));
        markdown.AppendLine();
        string fence = BuildFence(source);
        markdown.Append(fence).AppendLine(cellType == "code" ? language : "text");
        markdown.AppendLine(source.TrimEnd());
        markdown.AppendLine(fence);
        if (outputs.Count > 0) {
            markdown.AppendLine().AppendLine("#### Output").AppendLine();
            for (int index = 0; index < outputs.Count; index++) {
                NotebookOutput output = outputs[index];
                if (output.IsMarkdown) {
                    markdown.AppendLine(output.Text.Trim());
                } else {
                    string outputFence = BuildFence(output.Text);
                    markdown.AppendLine(outputFence).AppendLine(output.Text.TrimEnd()).AppendLine(outputFence);
                }
                if (index + 1 < outputs.Count) markdown.AppendLine();
            }
        }
        return markdown.ToString().TrimEnd();
    }

    private static string BuildPlainText(string source, IReadOnlyList<NotebookOutput> outputs) {
        return string.Join(
            Environment.NewLine + Environment.NewLine,
            new[] { source }.Concat(outputs.Select(static output => output.Text))
                .Where(static value => !string.IsNullOrWhiteSpace(value)))
            .Trim();
    }

    private static string BuildFence(string value) {
        int longest = 0;
        int current = 0;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '`') {
                current++;
                longest = Math.Max(longest, current);
            } else {
                current = 0;
            }
        }
        return new string('`', Math.Max(3, longest + 1));
    }

    private static string FindLanguage(JsonElement root) {
        if (!root.TryGetProperty("metadata", out JsonElement metadata) || metadata.ValueKind != JsonValueKind.Object) {
            return string.Empty;
        }
        if (metadata.TryGetProperty("kernelspec", out JsonElement kernel) && kernel.ValueKind == JsonValueKind.Object) {
            string? language = GetString(kernel, "language");
            if (!string.IsNullOrWhiteSpace(language)) return NormalizeFenceLanguage(language!);
        }
        if (metadata.TryGetProperty("language_info", out JsonElement languageInfo) && languageInfo.ValueKind == JsonValueKind.Object) {
            string? language = GetString(languageInfo, "name");
            if (!string.IsNullOrWhiteSpace(language)) return NormalizeFenceLanguage(language!);
        }
        return string.Empty;
    }

    private static string NormalizeFenceLanguage(string value) {
        return new string(value.Trim().Where(static character =>
            char.IsLetterOrDigit(character) || character is '-' or '_' or '+' or '.').ToArray());
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildMetadata(
        JsonElement root,
        string language,
        int totalCells,
        int includedCells) {
        if (root.TryGetProperty("nbformat", out JsonElement format) && format.TryGetInt32(out int major)) {
            yield return Metadata("notebook-nbformat", "NbFormat", major, "number");
        }
        if (root.TryGetProperty("nbformat_minor", out JsonElement minorFormat) && minorFormat.TryGetInt32(out int minor)) {
            yield return Metadata("notebook-nbformat-minor", "NbFormatMinor", minor, "number");
        }
        if (!string.IsNullOrWhiteSpace(language)) yield return Metadata("notebook-language", "Language", language, "string");
        yield return Metadata("notebook-cell-count", "CellCount", totalCells, "count");
        yield return Metadata("notebook-included-cell-count", "IncludedCellCount", includedCells, "count");
    }

    private static OfficeDocumentMetadataEntry Metadata(string id, string name, object value, string valueType) {
        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "notebook.document",
            Name = name,
            Value = Convert.ToString(value, CultureInfo.InvariantCulture),
            ValueType = valueType
        };
    }

    private static string? GetString(JsonElement owner, string propertyName) {
        return owner.TryGetProperty(propertyName, out JsonElement value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
    }

    private static string GetMultilineText(JsonElement owner, string propertyName) {
        return owner.TryGetProperty(propertyName, out JsonElement value) ? GetMultilineText(value) : string.Empty;
    }

    private static string GetMultilineText(JsonElement value) {
        if (value.ValueKind == JsonValueKind.String) return value.GetString() ?? string.Empty;
        if (value.ValueKind != JsonValueKind.Array) return string.Empty;
        var builder = new StringBuilder();
        foreach (JsonElement part in value.EnumerateArray()) {
            if (part.ValueKind == JsonValueKind.String) builder.Append(part.GetString());
        }
        return builder.ToString();
    }

    private static string Truncate(string value, int limit, string warning, List<string> warnings) {
        if (value.Length <= limit) return value;
        warnings.Add(warning);
        return value.Substring(0, limit);
    }

    private readonly struct NotebookOutput {
        internal NotebookOutput(string text, bool isMarkdown) {
            Text = text;
            IsMarkdown = isMarkdown;
        }

        internal string Text { get; }

        internal bool IsMarkdown { get; }
    }
}
