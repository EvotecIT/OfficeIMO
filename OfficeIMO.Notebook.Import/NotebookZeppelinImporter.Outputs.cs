using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookZeppelinImporter {
    private const int MaxTableProjectionCharacters = 2 * 1024 * 1024;
    private const int MaxTableColumns = 256;
    private const int MaxTableCells = 8_000;
    private static readonly char[] TableDelimiter = { '\t' };

    private static void ReadResults(JsonElement paragraph, NotebookCell sourceCell, bool markdown, string location,
        NotebookImportContext context, ref long totalOutputs) {
        JsonElement results = NotebookJson.Property(paragraph, "results");
        JsonElement messages = NotebookJson.Property(results, "msg");
        string? failure = NotebookJson.String(paragraph, "errorMessage");
        string? status = NotebookJson.String(paragraph, "status");
        bool failed = status == "ERROR" || NotebookJson.String(results, "code") == "ERROR";
        NotebookCell? outputCell = markdown ? null : sourceCell;
        if (messages.ValueKind == JsonValueKind.Array) {
            int count = messages.GetArrayLength();
            if (count > context.Limits.MaxOutputsPerCell) throw NotebookParser.Limit(nameof(context.Limits.MaxOutputsPerCell), context.Limits.MaxOutputsPerCell);
            totalOutputs += count;
            if (totalOutputs > context.Limits.MaxTotalOutputs) throw NotebookParser.Limit(nameof(context.Limits.MaxTotalOutputs), context.Limits.MaxTotalOutputs);
            int index = 0;
            foreach (JsonElement saved in messages.EnumerateArray()) {
                context.CancellationToken.ThrowIfCancellationRequested();
                string outputLocation = location + ".results.msg[" + (index++).ToString(CultureInfo.InvariantCulture) + "]";
                string? type = NotebookJson.String(saved, "type");
                string? payload = NotebookJson.String(saved, "data");
                if (type == "NULL" || (markdown && !failed && type == "HTML" &&
                    payload?.StartsWith("<div class=\"markdown-body\">", StringComparison.Ordinal) == true)) {
                    // A successful %md result is a renderer cache of the paragraph source.
                    continue;
                }
                NotebookOutput? output = payload == null ? null : ReadMessage(type, payload, failed, outputLocation, context);
                if (output != null) {
                    outputCell ??= context.AddCell(NotebookCellKind.Code, string.Empty);
                    context.AddOutput(outputCell, output);
                } else {
                    context.Report("notebook-zeppelin-result", "The saved result is malformed or has no supported static representation.", outputLocation);
                }
            }
        } else if (results.ValueKind is not (JsonValueKind.Undefined or JsonValueKind.Null) ||
            NotebookJson.Property(paragraph, "result").ValueKind != JsonValueKind.Undefined) {
            context.Report("notebook-zeppelin-result-list", "The paragraph uses an unsupported saved result structure.", location + ".results");
        }
        if (failed && !string.IsNullOrEmpty(failure) && !(outputCell?.Outputs.Any(static output => output.Kind == NotebookOutputKind.Error) ?? false)) {
            totalOutputs++;
            if (totalOutputs > context.Limits.MaxTotalOutputs) throw NotebookParser.Limit(nameof(context.Limits.MaxTotalOutputs), context.Limits.MaxTotalOutputs);
            outputCell ??= context.AddCell(NotebookCellKind.Code, string.Empty);
            context.AddOutput(outputCell, NotebookOutput.CreateError("ZeppelinError", failure!, Array.Empty<string>()));
        }
    }

    private static NotebookOutput? ReadMessage(string? type, string payload, bool failed, string location, NotebookImportContext context) {
        switch (type) {
            case "TEXT":
                return failed
                    ? NotebookOutput.CreateError("ZeppelinError", payload, Array.Empty<string>())
                    : NotebookOutput.CreateStream("stdout", payload);
            case "HTML":
                return Display("text/html", payload);
            case "SVG":
                return Display("image/svg+xml", payload);
            case "IMG":
                return ReadImage(payload, context.Limits.MaxInputBytes);
            case "TABLE":
                if (TryFormatTable(payload, context.CancellationToken, out string html)) {
                    return NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement>(StringComparer.Ordinal) {
                        ["text/html"] = NotebookImportContext.JsonString(html),
                        ["text/plain"] = NotebookImportContext.JsonString(payload)
                    });
                }
                context.Report("notebook-zeppelin-table", "The saved table is irregular or exceeds the static table limits. Tab-separated text is retained.",
                    location, OfficeConversionLossKind.Approximation);
                return Display("text/plain", payload);
            default:
                // Interactive Angular and network results have no static meaning. They never
                // become executable HTML or requests for external resources.
                return null;
        }
    }

    private static NotebookOutput Display(string mime, string value) =>
        NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement>(StringComparer.Ordinal) {
            [mime] = NotebookImportContext.JsonString(value)
        });

    private static NotebookOutput? ReadImage(string payload, long maximumImageBytes) {
        if (payload.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
            return OfficeImageReader.TryReadBase64DataUri(payload, maximumImageBytes, out byte[] image, out OfficeImageInfo identified)
                ? Display(identified.MimeType, Convert.ToBase64String(image))
                : null;
        }
        // Built-in Zeppelin renderers store base64 image bytes with no MIME declaration.
        byte[] bytes;
        try {
            bytes = Convert.FromBase64String(payload);
        } catch (FormatException) {
            return null;
        }
        return OfficeImageReader.TryIdentifyByContent(bytes, fileName: null, out OfficeImageInfo info) &&
            info.MimeType is "image/png" or "image/jpeg"
            ? Display(info.MimeType, payload)
            : null;
    }

    private static bool TryFormatTable(string table, CancellationToken token, out string html) {
        html = string.Empty;
        if (table.Length == 0 || table.Length > MaxTableProjectionCharacters / 2) return false;
        using var rows = new StringReader(table);
        string[] headers = rows.ReadLine()!.Split(TableDelimiter, MaxTableColumns + 1, StringSplitOptions.None);
        if (headers.Length > MaxTableColumns) return false;
        var output = new StringBuilder("<table><thead>");
        AppendTableRow(output, headers, "th");
        output.Append("</thead><tbody>");
        int count = headers.Length;
        string? line;
        while ((line = rows.ReadLine()) != null) {
            token.ThrowIfCancellationRequested();
            string[] row = line.Split(TableDelimiter, headers.Length + 1, StringSplitOptions.None);
            count += row.Length;
            if (row.Length != headers.Length || count > MaxTableCells) return false;
            AppendTableRow(output, row, "td");
            if (output.Length > MaxTableProjectionCharacters) return false;
        }
        output.Append("</tbody></table>");
        if (output.Length > MaxTableProjectionCharacters) return false;
        html = output.ToString();
        return true;
    }

    private static void AppendTableRow(StringBuilder output, string[] row, string tag) {
        output.Append("<tr>");
        foreach (string cell in row) {
            output.Append('<').Append(tag).Append('>').Append(WebUtility.HtmlEncode(cell))
                .Append("</").Append(tag).Append('>');
        }
        output.Append("</tr>");
    }
}
