using System.Globalization;
using System.Text.Json;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookZeppelinImporter {
    internal static void Read(ReadOnlyMemory<byte> input, NotebookImportContext context) {
        using JsonDocument json = JsonDocument.Parse(input, new JsonDocumentOptions { MaxDepth = context.Limits.MaxJsonDepth });
        JsonElement root = json.RootElement;
        string? title = NotebookJson.String(root, "name");
        string? noteId = NotebookJson.String(root, "id");
        JsonElement paragraphs = NotebookJson.Property(root, "paragraphs");
        if (string.IsNullOrWhiteSpace(noteId) || title == null || paragraphs.ValueKind != JsonValueKind.Array) {
            throw new FormatException("Expected a Zeppelin note export with an ID, name and paragraphs array.");
        }
        if (paragraphs.GetArrayLength() > context.Limits.MaxCells) throw NotebookParser.Limit(nameof(context.Limits.MaxCells), context.Limits.MaxCells);

        context.Document.SetMetadata("title", NotebookImportContext.JsonString(title));
        context.Document.SetMetadata("officeimo_zeppelin_note_id", NotebookImportContext.JsonString(noteId!));
        string? defaultInterpreter = NotebookJson.String(root, "defaultInterpreterGroup");
        if (IsNonemptyObject(root, "noteForms") || IsNonemptyObject(root, "angularObjects")) {
            context.Report("notebook-zeppelin-forms", "Zeppelin forms and Angular objects were omitted.", "noteForms");
        }

        long totalOutputs = 0;
        int index = 0;
        foreach (JsonElement paragraph in paragraphs.EnumerateArray()) {
            context.CancellationToken.ThrowIfCancellationRequested();
            string location = "paragraphs[" + (index++).ToString(CultureInfo.InvariantCulture) + "]";
            if (paragraph.ValueKind != JsonValueKind.Object) throw new FormatException("Each Zeppelin paragraph must be an object.");
            ReadParagraph(paragraph, defaultInterpreter, location, context, ref totalOutputs);
        }
    }

    private static void ReadParagraph(JsonElement paragraph, string? defaultInterpreter, string location,
        NotebookImportContext context, ref long totalOutputs) {
        string? id = NotebookJson.String(paragraph, "id");
        JsonElement script = NotebookJson.Property(paragraph, "text");
        if (script.ValueKind is not (JsonValueKind.Undefined or JsonValueKind.Null or JsonValueKind.String)) {
            context.Report("notebook-zeppelin-paragraph-text", "The paragraph source is not text.", location + ".text");
        }
        string source = script.ValueKind == JsonValueKind.String ? script.GetString()! : string.Empty;
        string? interpreter = ReadInterpreter(ref source);
        string? language = GetLanguage(interpreter ?? defaultInterpreter);
        bool markdown = language == "markdown";
        NotebookCell cell = context.AddCell(markdown ? NotebookCellKind.Markdown : NotebookCellKind.Code,
            source, markdown ? null : language, id);
        if (interpreter != null) cell.SetMetadata("officeimo_zeppelin_interpreter", NotebookImportContext.JsonString(interpreter));
        if (NotebookJson.String(paragraph, "title") is { Length: > 0 } paragraphTitle) {
            cell.SetMetadata("officeimo_zeppelin_paragraph_title", NotebookImportContext.JsonString(paragraphTitle));
        }

        JsonElement config = NotebookJson.Property(paragraph, "config");
        string? graph = NotebookJson.String(NotebookJson.Property(config, "graph"), "mode");
        if (!string.IsNullOrWhiteSpace(graph) && graph != "table") {
            context.Report("notebook-zeppelin-graph", "The Zeppelin graph configuration was omitted. Any saved table data is retained.",
                location + ".config.graph", OfficeConversionLossKind.Approximation);
        }
        ReadResults(paragraph, cell, markdown, location, context, ref totalOutputs);
    }

    private static string? ReadInterpreter(ref string source) {
        if (source.Length < 2 || source[0] != '%') return null;
        int end = 1;
        while (end < source.Length && !char.IsWhiteSpace(source[end])) {
            char part = source[end];
            if (!IsInterpreterCharacter(part)) return null;
            end++;
        }
        if (end == 1 || end > 65 || source[1] is '.' or '-' or '_') return null;
        string interpreter = source.Substring(1, end - 1);
        while (end < source.Length && source[end] is ' ' or '\t') end++;
        if (end < source.Length && source[end] == '\r') end++;
        if (end < source.Length && source[end] == '\n') end++;
        source = source.Substring(end);
        return interpreter;
    }

    private static bool IsInterpreterCharacter(char part) =>
        part is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9' or '.' or '_' or '-';

    private static string? GetLanguage(string? interpreter) => interpreter switch {
        "md" or "markdown" => "markdown",
        "sh" => "bash",
        "spark" or "livy.spark" => "scala",
        "python" or "python.ipython" or "spark.pyspark" or "pyspark" or "ipython" => "python",
        "sql" or "spark.sql" or "flink.ssql" or "flink.bsql" => "sql",
        "spark.r" or "sparkr" => "r",
        _ => interpreter
    };

    private static bool IsNonemptyObject(JsonElement owner, string key) =>
        NotebookJson.Property(owner, key) is { ValueKind: JsonValueKind.Object } value && value.EnumerateObject().MoveNext();
}
