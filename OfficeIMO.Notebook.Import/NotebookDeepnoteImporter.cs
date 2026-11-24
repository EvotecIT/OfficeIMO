using System.Globalization;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Yaml;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookDeepnoteImporter {
    private static readonly MarkdownSaveOptions LiteralTextWriter = new() {
        TextEscapingMode = MarkdownTextEscapingMode.LiteralBackslash,
        OutputLineEnding = "\n"
    };

    internal static void Read(ReadOnlyMemory<byte> input, NotebookImportContext context) {
        using YamlDocument snapshot = Parse(input, context);
        if (snapshot.Documents.Count != 1 || snapshot.Root is not YamlMapping root ||
            Text(root, "version") != "1.0.0" || Get(root, "project") is not YamlMapping project ||
            Get(project, "notebooks") is not YamlSequence notebooks || notebooks.Count == 0) {
            throw new FormatException("Expected a Deepnote 1.0.0 document containing project.notebooks.");
        }

        // The producer emits a JSON-like YAML tree. Aliases are outside this profile: converting
        // a large anchored scalar to JSON once per reference could otherwise amplify the input bytes.
        CheckTree(root, context.CancellationToken);
        int selectedIndex = SelectNotebook(notebooks, context.Options.DeepnoteNotebookId, context.CancellationToken);
        var notebook = (YamlMapping)notebooks[selectedIndex];
        string location = "project.notebooks[" + selectedIndex.ToString(CultureInfo.InvariantCulture) + "]";
        if (Get(notebook, "blocks") is not YamlSequence blocks) {
            throw new FormatException("The selected Deepnote notebook must contain a blocks list.");
        }
        if (blocks.Count > context.Limits.MaxCells) throw NotebookParser.Limit(nameof(context.Limits.MaxCells), context.Limits.MaxCells);

        SetMetadata(context.Document, "title", Text(notebook, "name"));
        SetMetadata(context.Document, "officeimo_deepnote_notebook_id", Text(notebook, "id"));
        SetMetadata(context.Document, "officeimo_deepnote_project_name", Text(project, "name"));
        if (HasContent(root, "environment") || HasContent(root, "integrations") ||
            HasContent(project, "integrations") || HasContent(project, "settings")) {
            context.Report("notebook-deepnote-project-settings",
                "Deepnote environments, integrations and project settings were omitted.", "project");
        }

        var ordered = new List<(YamlMapping Block, string SortKey, int Index)>(blocks.Count);
        for (int i = 0; i < blocks.Count; i++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (blocks[i] is not YamlMapping block || Text(block, "sortingKey") is not { } sortingKey ||
                string.IsNullOrEmpty(Text(block, "type"))) {
                throw new FormatException("Each Deepnote block must contain its type and sorting key.");
            }
            ordered.Add((block, sortingKey, i));
        }
        long outputCount = 0;
        // The Deepnote comparison uses ordinal UTF-16 code units and keeps source order for ties.
        foreach (var entry in ordered.OrderBy(static entry => entry.SortKey, StringComparer.Ordinal)) {
            string blockLocation = location + ".blocks[" + entry.Index.ToString(CultureInfo.InvariantCulture) + "]";
            ReadBlock(entry.Block, blockLocation, context, ref outputCount);
        }
    }

    private static YamlDocument Parse(ReadOnlyMemory<byte> input, NotebookImportContext context) {
        string text = NotebookImportContext.DecodeUtf8(input);
        YamlLoadOptions options = YamlLoadOptions.UntrustedDefaults;
        options.Strict = true;
        options.Schema = YamlSchema.Core;
        options.MaxDepth = Math.Min(context.Limits.MaxJsonDepth, 512);
        options.MaxInputBytes = context.Limits.MaxInputBytes;
        options.MaxScalarLength = (int)Math.Min(context.Limits.MaxInputBytes, int.MaxValue);
        try {
            return YamlDocument.Parse(text, options, context.CancellationToken);
        } catch (YamlLimitExceededException exception) {
            throw new InvalidDataException("The Deepnote YAML exceeds an import limit.", exception);
        } catch (YamlException exception) {
            throw new FormatException("The Deepnote YAML is invalid.", exception);
        }
    }

    private static int SelectNotebook(YamlSequence notebooks, string? selectedId, CancellationToken token) {
        if (selectedId == null && notebooks.Count > 1) {
            throw new FormatException("This Deepnote project contains multiple notebooks. Specify DeepnoteNotebookId.");
        }
        int selected = -1;
        for (int i = 0; i < notebooks.Count; i++) {
            token.ThrowIfCancellationRequested();
            if (notebooks[i] is not YamlMapping item || string.IsNullOrWhiteSpace(Text(item, "id"))) {
                throw new FormatException("Each Deepnote notebook must contain its ID.");
            }
            if (selectedId != null && !string.Equals(Text(item, "id"), selectedId, StringComparison.Ordinal)) continue;
            if (selected >= 0) throw new FormatException("The selected Deepnote notebook ID is duplicated.");
            selected = i;
        }
        if (selected < 0) throw new FormatException("The selected Deepnote notebook ID was not found.");
        return selected;
    }

    private static void ReadBlock(YamlMapping block, string location, NotebookImportContext context, ref long totalOutputs) {
        string type = Text(block, "type")!;
        string? id = Text(block, "id");
        YamlNode? content = Get(block, "content");
        string source = Text(block, "content") ?? string.Empty;
        YamlNode? outputNode = Get(block, "outputs");
        var outputs = outputNode as YamlSequence;
        if (content is not (null or YamlScalar)) context.Report("notebook-deepnote-block-content", "The block content is not text.", location + ".content");
        if (outputNode is not (null or YamlSequence)) context.Report("notebook-deepnote-output-list", "The saved outputs are not a list.", location + ".outputs");
        if (outputs != null) {
            if (outputs.Count > context.Limits.MaxOutputsPerCell) throw NotebookParser.Limit(nameof(context.Limits.MaxOutputsPerCell), context.Limits.MaxOutputsPerCell);
            totalOutputs += outputs.Count;
            if (totalOutputs > context.Limits.MaxTotalOutputs) throw NotebookParser.Limit(nameof(context.Limits.MaxTotalOutputs), context.Limits.MaxTotalOutputs);
        }

        NotebookCell cell;
        switch (type) {
            case "code":
            case "sql":
                cell = context.AddCell(NotebookCellKind.Code, source, type == "sql" ? "sql" : "python", id);
                break;
            case "image":
                cell = AddImage(block, location, id, context);
                break;
            case "markdown":
            case "separator":
            case "text-cell-h1":
            case "text-cell-h2":
            case "text-cell-h3":
            case "text-cell-p":
            case "text-cell-bullet":
            case "text-cell-todo":
            case "text-cell-callout":
                cell = context.AddCell(NotebookCellKind.Markdown,
                    RenderTextBlock(type, source, Get(block, "metadata") as YamlMapping, context.CancellationToken), sourceId: id);
                if (type == "text-cell-callout") {
                    context.Report("notebook-deepnote-callout", "The Deepnote callout is represented as a block quote.", location, OfficeConversionLossKind.Approximation);
                }
                break;
            default:
                // Jupyter raw cells cannot carry saved outputs. Use an empty code cell when output
                // exists so a vendor prompt, input, or configuration never becomes executable source.
                bool hasSavedOutput = outputs is { Count: > 0 };
                cell = context.AddCell(hasSavedOutput ? NotebookCellKind.Code : NotebookCellKind.Raw,
                    hasSavedOutput ? string.Empty : source, sourceId: id);
                cell.SetMetadata("officeimo_deepnote_block_type", NotebookImportContext.JsonString(type));
                if (hasSavedOutput && source.Length > 0) {
                    cell.SetMetadata("officeimo_deepnote_block_source", NotebookImportContext.JsonString(source));
                }
                context.Report("notebook-deepnote-block-type",
                    "This Deepnote block has no portable presentation. Its source and supported saved outputs are retained.",
                    location, OfficeConversionLossKind.Approximation);
                break;
        }

        if (Text(block, "contentHash") is { } hash) cell.SetMetadata("officeimo_deepnote_content_hash", NotebookImportContext.JsonString(hash));
        if (cell.Kind == NotebookCellKind.Code) {
            cell.ExecutionCount = Count(block, "executionCount", location, context);
            if (outputs != null) ReadOutputs(outputs, cell, location, context);
        } else if (outputs is { Count: > 0 }) {
            context.Report("notebook-deepnote-noncode-outputs", "Saved outputs on this non-code block were omitted.", location + ".outputs");
        }
    }

    private static NotebookCell AddImage(YamlMapping block, string location, string? id, NotebookImportContext context) {
        string? path = Text(Get(block, "metadata") as YamlMapping, "deepnote_img_src");
        if (!OfficeImageReader.TryReadBase64DataUri(path, context.Limits.MaxInputBytes, out byte[] bytes, out OfficeImageInfo info)) {
            context.Report("notebook-deepnote-image", "The image block has no supported embedded image. External references are not opened.", location);
            return context.AddCell(NotebookCellKind.Raw, string.Empty, sourceId: id);
        }
        string attachment = "deepnote-image" + OfficeImageInfo.GetDefaultExtension(info.Format);
        NotebookCell image = context.AddCell(NotebookCellKind.Markdown, "![Deepnote image](attachment:" + attachment + ")", sourceId: id);
        image.SetAttachment(attachment, new Dictionary<string, JsonElement>(StringComparer.Ordinal) {
            [info.MimeType] = NotebookImportContext.JsonString(Convert.ToBase64String(bytes))
        });
        return image;
    }

    private static string RenderTextBlock(string type, string source, YamlMapping? metadata, CancellationToken token) {
        if (type == "markdown") return source;
        if (type == "separator") return "---";
        using MarkdownDocument markdown = MarkdownDocument.Create();
        switch (type) {
            case "text-cell-h1":
            case "text-cell-h2":
            case "text-cell-h3":
                markdown.AddHeading(source, type[type.Length - 1] - '0');
                break;
            case "text-cell-bullet":
            case "text-cell-todo":
                var list = new UnorderedListBlock();
                bool isChecked = Get(metadata, "checked") is YamlScalar flag && flag.TryGetValue(out bool value) && value;
                list.Items.Add(type == "text-cell-todo" ? ListItem.Task(source, isChecked) : ListItem.Text(source));
                markdown.Add(list);
                break;
            case "text-cell-callout":
                var quote = new QuoteBlock();
                quote.ChildBlocks.Add(new ParagraphBlock(new InlineSequence().Text(source)));
                markdown.Add(quote);
                break;
            default:
                markdown.AddParagraph(source);
                break;
        }
        return markdown.ToMarkdown(LiteralTextWriter, token).TrimEnd('\n');
    }

    private static void CheckTree(YamlNode node, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (node) {
            case YamlAlias:
                throw new FormatException("YAML aliases are outside the Deepnote import profile.");
            case YamlMapping mapping:
                foreach (YamlMappingEntry entry in mapping) {
                    if (entry.Key is not YamlScalar) throw new FormatException("Deepnote mapping keys must be scalars.");
                    CheckTree(entry.Value, cancellationToken);
                }
                break;
            case YamlSequence sequence:
                foreach (YamlNode item in sequence) CheckTree(item, cancellationToken);
                break;
        }
    }

    private static YamlNode? Get(YamlMapping? mapping, string name) =>
        mapping != null && mapping.TryGetValue(name, out YamlNode? value) ? value : null;

    private static string? Text(YamlMapping? mapping, string name) =>
        Get(mapping, name) is YamlScalar { IsNull: false } scalar ? scalar.AsString() : null;

    private static bool HasContent(YamlMapping mapping, string name) => Get(mapping, name) switch {
        null or YamlScalar { IsNull: true } => false,
        YamlMapping child => child.Count != 0,
        YamlSequence child => child.Count != 0,
        _ => true
    };

    private static void SetMetadata(NotebookDocument notebook, string name, string? value) {
        if (value != null) notebook.SetMetadata(name, NotebookImportContext.JsonString(value));
    }
}
