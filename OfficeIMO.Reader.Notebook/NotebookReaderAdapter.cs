using global::OfficeIMO.Notebook;

namespace OfficeIMO.Reader.Notebook;

internal static class NotebookReaderAdapter {
    private const int MaxFenceLanguageCharacters = 64;

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
        int outputLimit = notebookOptions.IncludeOutputs && notebookOptions.IncludeCodeCells ? notebookOptions.MaxOutputsPerCell : 0;
        using NotebookDocument notebook = NotebookDocument.LoadForReader(input.Bytes, notebookOptions.MaxCells, outputLimit, cancellationToken);
        string language = NormalizeFenceLanguage(notebook.Language ?? string.Empty);
        var chunks = new List<ReaderChunk>(Math.Min(notebook.Cells.Count, notebookOptions.MaxCells));
        var adapterDiagnostics = new List<OfficeDocumentDiagnostic>();
        int includedCells = 0;
        foreach (NotebookCell cell in notebook.Cells) {
            cancellationToken.ThrowIfCancellationRequested();
            int cellIndex = cell.SourceIndex;
            if (cell.Kind == NotebookCellKind.Unknown
                || (cell.Kind == NotebookCellKind.Code && !notebookOptions.IncludeCodeCells)) continue;
            string cellType = cell.Kind switch {
                NotebookCellKind.Code => "code",
                NotebookCellKind.Markdown => "markdown",
                _ => "raw"
            };

            var warnings = new List<string>();
            // Reader permits the complete first Unicode scalar at its one-character boundary.
            string source = cell.GetSource(Math.Max(2, notebookOptions.MaxCellCharacters), out bool truncated);
            if (truncated || source.Length > notebookOptions.MaxCellCharacters) {
                source = DocumentReaderEngine.TruncateAdapterProjection(source, notebookOptions.MaxCellCharacters);
                warnings.Add("Notebook cell source was truncated at MaxCellCharacters.");
            }
            IReadOnlyList<ProjectedOutput> outputs = cell.Kind == NotebookCellKind.Code && notebookOptions.IncludeOutputs
                ? ReadOutputs(cell, notebookOptions, warnings, cancellationToken)
                : Array.Empty<ProjectedOutput>();
            string markdown = BuildMarkdown(cellType, cellIndex, language, source, outputs);
            string text = BuildPlainText(source, outputs);
            if (string.IsNullOrWhiteSpace(markdown) && string.IsNullOrWhiteSpace(text)) {
                AddDroppedCellDiagnostics(
                    adapterDiagnostics,
                    warnings,
                    input.Source.Path,
                    cellIndex,
                    cellType,
                    notebookOptions.MaxCellCharacters);
                continue;
            }

            IReadOnlyList<string> textParts = DocumentReaderEngine.SplitAdapterProjection(text, readerOptions.MaxChars);
            IReadOnlyList<string> markdownParts = DocumentReaderEngine.SplitAdapterProjection(markdown, readerOptions.MaxChars);
            int partCount = Math.Max(textParts.Count, markdownParts.Count);
            bool wasSplit = partCount > 1;
            string[]? chunkWarnings = BuildChunkWarnings(warnings, wasSplit);
            string baseAnchor = "notebook-cell-" + cellIndex.ToString("D4", CultureInfo.InvariantCulture);
            for (int partIndex = 0; partIndex < partCount; partIndex++) {
                string anchor = partIndex == 0
                    ? baseAnchor
                    : baseAnchor + "-part-" + partIndex.ToString("D4", CultureInfo.InvariantCulture);
                var location = new ReaderLocation {
                    Path = input.Source.Path,
                    BlockIndex = chunks.Count,
                    SourceBlockIndex = cellIndex,
                    SourceBlockKind = "notebook-" + cellType + "-cell",
                    BlockAnchor = anchor
                };
                var chunk = new ReaderChunk {
                    Id = anchor,
                    Kind = ReaderInputKind.Json,
                    ContinuesPreviousChunk = partIndex > 0,
                    Location = location,
                    Text = partIndex < textParts.Count ? textParts[partIndex] : string.Empty,
                    Markdown = partIndex < markdownParts.Count ? markdownParts[partIndex] : string.Empty,
                    Warnings = chunkWarnings
                };
                DocumentReaderEngine.ApplyAdapterSource(chunk, input, readerOptions.ComputeHashes);
                chunks.Add(chunk);
            }
            includedCells++;
        }

        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Json,
            input.Source,
            new[] { OfficeDocumentReaderBuilderNotebookExtensions.HandlerId, "jupyter.nbformat" });
        result.Source.Title = Path.GetFileName(input.Source.Path ?? "notebook.ipynb");
        int totalCells = notebook.SourceCellCount;
        result.Metadata = result.Metadata.Concat(BuildMetadata(notebook, language, totalCells, includedCells)).ToArray();
        if (totalCells > notebookOptions.MaxCells) {
            adapterDiagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = "notebook-cell-limit",
                Message = "Notebook cell inspection stopped at MaxCells.",
                Source = OfficeDocumentReaderBuilderNotebookExtensions.HandlerId,
                IsRecoverable = true,
                Attributes = new Dictionary<string, string> {
                    ["maxCells"] = notebookOptions.MaxCells.ToString(CultureInfo.InvariantCulture),
                    ["totalCells"] = totalCells.ToString(CultureInfo.InvariantCulture)
                }
            });
        }
        if (adapterDiagnostics.Count > 0) {
            result.Diagnostics = result.Diagnostics.Concat(adapterDiagnostics).ToArray();
        }
        return result;
    }

    private static void AddDroppedCellDiagnostics(
        List<OfficeDocumentDiagnostic> diagnostics,
        IReadOnlyList<string> warnings,
        string? path,
        int cellIndex,
        string cellType,
        int maxCellCharacters) {
        string anchor = "notebook-cell-" + cellIndex.ToString("D4", CultureInfo.InvariantCulture);
        for (int warningIndex = 0; warningIndex < warnings.Count; warningIndex++) {
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "notebook-content-truncated",
                Message = warnings[warningIndex],
                Source = OfficeDocumentReaderBuilderNotebookExtensions.HandlerId,
                IsRecoverable = true,
                Location = new ReaderLocation {
                    Path = path,
                    SourceBlockIndex = cellIndex,
                    SourceBlockKind = "notebook-" + cellType + "-cell",
                    BlockAnchor = anchor
                },
                Attributes = new Dictionary<string, string> {
                    ["cellIndex"] = cellIndex.ToString(CultureInfo.InvariantCulture),
                    ["cellType"] = cellType,
                    ["maxCellCharacters"] = maxCellCharacters.ToString(CultureInfo.InvariantCulture)
                }
            });
        }
    }

    private static IReadOnlyList<ProjectedOutput> ReadOutputs(
        NotebookCell cell,
        ReaderNotebookOptions options,
        List<string> warnings,
        CancellationToken cancellationToken) {
        int sourceOutputCount = cell.SourceOutputCount;
        if (sourceOutputCount == 0) return Array.Empty<ProjectedOutput>();
        var projected = new List<ProjectedOutput>(cell.Outputs.Count);
        int totalCharacters = 0;
        foreach (NotebookOutput output in cell.Outputs) {
            cancellationToken.ThrowIfCancellationRequested();
            int remaining = options.MaxOutputCharactersPerCell - totalCharacters;
            ProjectedOutput? value = ProjectOutput(output, Math.Max(2, remaining), out bool truncated);
            if (!value.HasValue || (string.IsNullOrWhiteSpace(value.Value.Text) && !truncated)) continue;
            if (remaining <= 0) {
                warnings.Add("Notebook output text was truncated at MaxOutputCharactersPerCell.");
                return projected;
            }
            string text = value.Value.Text;
            if (truncated || text.Length > remaining) {
                text = DocumentReaderEngine.TruncateAdapterProjection(text, remaining);
                warnings.Add("Notebook output text was truncated at MaxOutputCharactersPerCell.");
            }
            projected.Add(new ProjectedOutput(text, value.Value.IsMarkdown));
            totalCharacters += text.Length;
        }
        if (sourceOutputCount > options.MaxOutputsPerCell) warnings.Add("Notebook outputs were truncated at MaxOutputsPerCell.");
        return projected;
    }

    private static ProjectedOutput? ProjectOutput(NotebookOutput output, int maxCharacters, out bool truncated) {
        truncated = false;
        if (output.Kind == NotebookOutputKind.Stream) {
            return new ProjectedOutput(output.GetText(maxCharacters, out truncated), false);
        }
        if (output.Kind == NotebookOutputKind.Error) {
            string traceback = string.Join("\n", output.Traceback);
            if (!string.IsNullOrWhiteSpace(traceback)) return new ProjectedOutput(traceback, false);
            string error = string.Join(": ", new[] { output.ErrorName, output.ErrorValue }
                .Where(static value => !string.IsNullOrWhiteSpace(value)));
            return new ProjectedOutput(error, false);
        }
        if (output.Data.TryGetValue("text/markdown", out JsonElement markdown)) {
            string markdownText = NotebookJson.ReadText(markdown, maxCharacters, out truncated);
            if (!string.IsNullOrWhiteSpace(markdownText) || truncated) return new ProjectedOutput(markdownText, true);
        }
        if (output.Data.TryGetValue("text/plain", out JsonElement plain)) {
            return new ProjectedOutput(NotebookJson.ReadText(plain, maxCharacters, out truncated), false);
        }
        return null;
    }

    private static string BuildMarkdown(
        string cellType,
        int cellIndex,
        string language,
        string source,
        IReadOnlyList<ProjectedOutput> outputs) {
        if (cellType == "markdown") return source.TrimEnd();
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
                ProjectedOutput output = outputs[index];
                if (output.IsMarkdown) {
                    markdown.AppendLine(output.Text.TrimEnd());
                } else {
                    string outputFence = BuildFence(output.Text);
                    markdown.AppendLine(outputFence).AppendLine(output.Text.TrimEnd()).AppendLine(outputFence);
                }
                if (index + 1 < outputs.Count) markdown.AppendLine();
            }
        }
        return markdown.ToString().TrimEnd();
    }

    private static string BuildPlainText(string source, IReadOnlyList<ProjectedOutput> outputs) {
        return string.Join(
            Environment.NewLine + Environment.NewLine,
            new[] { source }.Concat(outputs.Select(static output => output.Text))
                .Where(static value => !string.IsNullOrWhiteSpace(value)))
            .TrimEnd();
    }

    private static string[]? BuildChunkWarnings(IReadOnlyList<string> warnings, bool wasSplit) {
        if (!wasSplit) return warnings.Count == 0 ? null : warnings.ToArray();
        return warnings
            .Concat(new[] { "Notebook content was split due to MaxChars." })
            .Distinct(StringComparer.Ordinal)
            .ToArray();
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

    private static string NormalizeFenceLanguage(string value) {
        var normalized = new StringBuilder(Math.Min(value.Length, MaxFenceLanguageCharacters));
        for (int index = 0; index < value.Length && normalized.Length < MaxFenceLanguageCharacters; index++) {
            char character = value[index];
            if (char.IsLetterOrDigit(character) || character is '-' or '_' or '+' or '.') {
                normalized.Append(character);
            }
        }
        return normalized.ToString();
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildMetadata(
        NotebookDocument notebook,
        string language,
        int totalCells,
        int includedCells) {
        if (notebook.DeclaredNbFormat is int major) {
            yield return Metadata("notebook-nbformat", "NbFormat", major, "number");
        }
        if (notebook.DeclaredNbFormatMinor is int minor) {
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

    private readonly struct ProjectedOutput {
        internal ProjectedOutput(string text, bool isMarkdown) {
            Text = text;
            IsMarkdown = isMarkdown;
        }

        internal string Text { get; }

        internal bool IsMarkdown { get; }
    }
}
