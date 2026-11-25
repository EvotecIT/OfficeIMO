namespace OfficeIMO.Notebook;

internal static class NotebookWriter {
    internal static void WriteCell(Utf8JsonWriter writer, NotebookCell cell, bool legacy, CancellationToken cancellationToken) {
        if (!legacy && cell.Kind == NotebookCellKind.Unknown && !cell.SourceEdited && !cell.MetadataEdited &&
            NotebookJson.String(cell.Original, "id") == cell.Id) {
            cell.Original.WriteTo(writer);
            return;
        }

        writer.WriteStartObject();
        if (cell.Kind == NotebookCellKind.Unknown) {
            NotebookJson.CopyProperties(writer, cell.Original, static property => property.Name is not ("id" or "source" or "metadata"));
        } else {
            NotebookJson.CopyProperties(writer, cell.Original, property => IncludeAdditionalCellProperty(property, cell, legacy));
            string kind = cell.Kind == NotebookCellKind.Markdown ? "markdown" : cell.OriginalKind;
            writer.WriteString("cell_type", kind);
        }
        writer.WriteString("id", cell.Id);
        if (legacy || cell.MetadataEdited ||
            !NotebookJson.WriteOriginal(writer, "metadata", NotebookJson.Property(cell.Original, "metadata"))) {
            WriteCellMetadata(writer, cell, legacy);
        }
        bool keepSource = !legacy && !cell.SourceEdited && NotebookJson.WriteOriginal(writer, "source", cell.SourceElement);
        if (!keepSource && (cell.Kind != NotebookCellKind.Unknown || cell.SourceEdited)) writer.WriteString("source", cell.Source);

        if (cell.Kind == NotebookCellKind.Code) {
            WriteCodeCell(writer, cell, legacy, cancellationToken);
        } else if (cell.Kind is NotebookCellKind.Markdown or NotebookCellKind.Raw) {
            WriteAttachments(writer, cell, legacy);
        }
        writer.WriteEndObject();
    }

    private static void WriteCellMetadata(Utf8JsonWriter writer, NotebookCell cell, bool legacy) {
        IReadOnlyDictionary<string, JsonElement> metadata = cell.Metadata;
        if (legacy && cell.Kind == NotebookCellKind.Code &&
            NotebookJson.Property(cell.Original, "language") is { ValueKind: JsonValueKind.String } sourceLanguage &&
            !string.IsNullOrWhiteSpace(sourceLanguage.GetString()) &&
            (!metadata.TryGetValue("language", out JsonElement currentLanguage) || !JsonElement.DeepEquals(sourceLanguage, currentLanguage))) {
            // Version 4 carries the legacy source language in metadata.
            var normalized = metadata.ToDictionary(static property => property.Key, static property => property.Value, StringComparer.Ordinal);
            normalized["language"] = sourceLanguage;
            metadata = normalized;
        }
        NotebookJson.WriteMap(writer, "metadata", metadata);
    }

    private static void WriteCodeCell(Utf8JsonWriter writer, NotebookCell cell, bool legacy, CancellationToken cancellationToken) {
        if (legacy || cell.ExecutionCountEdited ||
            !NotebookJson.WriteOriginal(writer, "execution_count", NotebookJson.Property(cell.Original, "execution_count"))) {
            WriteCount(writer, cell.ExecutionCount);
        }
        JsonElement original = NotebookJson.Property(cell.Original, "outputs");
        if (!legacy && !cell.OutputsEdited && original.ValueKind != JsonValueKind.Array &&
            NotebookJson.WriteOriginal(writer, "outputs", original)) return;
        writer.WriteStartArray("outputs");
        foreach (NotebookOutput output in cell.Outputs) {
            cancellationToken.ThrowIfCancellationRequested();
            WriteOutput(writer, output);
        }
        writer.WriteEndArray();
    }

    private static void WriteAttachments(Utf8JsonWriter writer, NotebookCell cell, bool legacy) {
        if (!legacy && !cell.AttachmentsEdited &&
            NotebookJson.WriteOriginal(writer, "attachments", NotebookJson.Property(cell.Original, "attachments"))) return;
        if (cell.Attachments.Count == 0) return;
        writer.WriteStartObject("attachments");
        foreach (KeyValuePair<string, IReadOnlyDictionary<string, JsonElement>> attachment in cell.Attachments) {
            NotebookJson.WriteMap(writer, attachment.Key, attachment.Value);
        }
        writer.WriteEndObject();
    }

    private static bool IncludeAdditionalCellProperty(JsonProperty property, NotebookCell cell, bool legacy) {
        if (property.Name is "id" or "cell_type" or "metadata" or "source") return false;
        // The Deepnote Jupyter exporter writes the same group ID outside and inside metadata.
        // The extra cell property is invalid in nbformat 4.5. Keep the canonical metadata value.
        if (property.NameEquals("block_group") && property.Value.ValueKind == JsonValueKind.String &&
            cell.Metadata.TryGetValue("deepnote_block_group", out JsonElement group) && group.ValueKind == JsonValueKind.String &&
            JsonElement.DeepEquals(property.Value, group)) return false;
        if (cell.Kind == NotebookCellKind.Code) {
            return property.Name is not ("execution_count" or "outputs") &&
                !(legacy && property.Name is "input" or "prompt_number" or "collapsed" or "language");
        }
        if (cell.Kind is NotebookCellKind.Markdown or NotebookCellKind.Raw && property.Name == "attachments") return false;
        return !(legacy && cell.OriginalKind == "heading" && property.Name == "level");
    }

    private static void WriteOutput(Utf8JsonWriter writer, NotebookOutput output) {
        if ((!output.Legacy || output.Kind == NotebookOutputKind.Unknown) &&
            output.Original.ValueKind != JsonValueKind.Undefined) {
            output.Original.WriteTo(writer);
            return;
        }

        writer.WriteStartObject();
        switch (output.Kind) {
            case NotebookOutputKind.Stream:
                NotebookJson.CopyProperties(writer, output.Original, static property =>
                    property.Name is not ("output_type" or "name" or "stream" or "text"));
                writer.WriteString("output_type", "stream");
                writer.WriteString("name", output.Name);
                writer.WriteString("text", output.Text);
                break;
            case NotebookOutputKind.DisplayData:
            case NotebookOutputKind.ExecuteResult:
                writer.WriteString("output_type", output.Kind == NotebookOutputKind.DisplayData ? "display_data" : "execute_result");
                if (output.Kind == NotebookOutputKind.ExecuteResult) WriteCount(writer, output.ExecutionCount);
                NotebookJson.WriteMap(writer, "data", output.Data);
                NotebookJson.WriteMap(writer, "metadata", output.Metadata);
                break;
            case NotebookOutputKind.Error:
                NotebookJson.CopyProperties(writer, output.Original, static property =>
                    property.Name is not ("output_type" or "ename" or "evalue" or "traceback"));
                writer.WriteString("output_type", "error");
                writer.WriteString("ename", output.ErrorName ?? string.Empty);
                writer.WriteString("evalue", output.ErrorValue ?? string.Empty);
                writer.WriteStartArray("traceback");
                foreach (string line in output.Traceback) writer.WriteStringValue(line);
                writer.WriteEndArray();
                break;
        }
        writer.WriteEndObject();
    }

    private static void WriteCount(Utf8JsonWriter writer, int? executionCount) {
        if (executionCount.HasValue) writer.WriteNumber("execution_count", executionCount.Value);
        else writer.WriteNull("execution_count");
    }
}
