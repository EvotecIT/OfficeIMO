using System.Globalization;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Notebook.Import;

internal static partial class RNotebookHtmlImporter {
    internal static void Read(ReadOnlyMemory<byte> input, NotebookImportContext context) {
        string source = NotebookImportContext.DecodeUtf8(input);
        HtmlConversionDocumentOptions settings = HtmlConversionDocumentOptions.UntrustedDefaults;
        settings.IncludeNormalizedHtml = false;
        settings.Limits.MaxInputCharacters = (int)context.Limits.MaxInputBytes;
        settings.Limits.MaxHtmlDepth = 128;
        HtmlDocument html = HtmlConversionDocument.Parse(source, settings, context.CancellationToken).Document;
        HtmlElement body = html.Body ?? throw new InvalidDataException("The R Notebook does not contain an HTML body.");
        string rMarkdown = ReadEmbeddedSource(body, context);
        NotebookMarkupImporter.ReadSource(rMarkdown, NotebookImportFormat.RNotebookHtml, context);

        var parsed = new ChunkReader(context);
        parsed.Visit(body);
        parsed.Complete();
        RestoreSavedOutputs(parsed.Chunks, context);
    }

    private static string ReadEmbeddedSource(HtmlElement body, NotebookImportContext context) {
        HtmlElement? marker = null;
        foreach (HtmlNode node in body.Descendants()) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (node is not HtmlElement { LocalName: "div", Id: "rmd-source-code" } element) continue;
            if (marker != null) throw new InvalidDataException("The R Notebook contains multiple embedded source documents.");
            marker = element;
        }
        if (marker == null || marker.ChildNodes.Any(static node => node.Kind != HtmlNodeKind.Text)) {
            throw new InvalidDataException("The HTML does not contain an embedded R Markdown notebook source.");
        }

        byte[] bytes;
        try {
            bytes = Convert.FromBase64String(marker.TextContent);
        } catch (FormatException error) {
            throw new InvalidDataException("The embedded R Notebook source is not valid base64.", error);
        }
        return NotebookImportContext.DecodeUtf8(bytes);
    }

    private static void RestoreSavedOutputs(IReadOnlyList<SavedChunk> chunks, NotebookImportContext context) {
        if (chunks.Count == 0) return;
        NotebookCell[] codeCells = context.Document.Cells.Where(static cell => cell.Kind == NotebookCellKind.Code).ToArray();
        string[] sourceKeys = codeCells.Select(static cell => CodeKey(cell.Source)).ToArray();
        string[] chunkKeys = chunks.Select(static chunk => CodeKey(chunk.Source.ToString())).ToArray();
        bool ordered = codeCells.Length == chunks.Count;
        for (int i = 0; ordered && i < chunkKeys.Length; i++) {
            ordered = chunkKeys[i].Length == 0 || chunkKeys[i] == sourceKeys[i];
        }

        // The producer normally writes one ordered chunk group per source chunk. If rendered
        // source contradicts that order, only uniquely identifiable source can establish a match.
        Dictionary<string, int>? uniqueSources = ordered ? null : UniqueKeys(sourceKeys);
        Dictionary<string, int>? uniqueChunks = ordered ? null : UniqueKeys(chunkKeys);
        for (int i = 0; i < chunks.Count; i++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            SavedChunk chunk = chunks[i];
            if (chunk.Outputs.Count == 0) continue;

            NotebookCell target;
            if (ordered) {
                target = codeCells[i];
            } else if (uniqueChunks!.ContainsKey(chunkKeys[i]) && uniqueSources!.TryGetValue(chunkKeys[i], out int index)) {
                target = codeCells[index];
            } else {
                context.Report("notebook-import-rnb-unmatched-output",
                    "Saved output could not be assigned to a source chunk with confidence. It was retained in a separate cell.",
                    ChunkLocation(i), OfficeConversionLossKind.Approximation);
                target = context.AddCell(NotebookCellKind.Code, string.Empty, "r");
            }

            foreach (NotebookOutput output in chunk.Outputs) context.AddOutput(target, output);
        }
    }

    private static Dictionary<string, int> UniqueKeys(string[] keys) {
        var unique = new Dictionary<string, int>(StringComparer.Ordinal);
        var duplicate = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < keys.Length; i++) {
            string key = keys[i];
            if (key.Length == 0 || duplicate.Contains(key)) continue;
            if (unique.ContainsKey(key)) {
                unique.Remove(key);
                duplicate.Add(key);
            } else {
                unique.Add(key, i);
            }
        }
        return unique;
    }

    private static string CodeKey(string code) => code.Replace("\r\n", "\n").Replace('\r', '\n').Trim();

    private static string ChunkLocation(int index) => $"rnb.chunks[{index.ToString(CultureInfo.InvariantCulture)}]";
}
