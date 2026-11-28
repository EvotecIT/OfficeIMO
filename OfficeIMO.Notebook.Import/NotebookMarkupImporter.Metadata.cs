using System.Text.Json;
using OfficeIMO.Yaml;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookMarkupImporter {
    private static void ReadFrontMatter(ref SourceCursor body, ref string? language, NotebookImportContext context) {
        SourceCursor header = body;
        if (!header.TryRead(out SourceLine opening) || !opening.Text(body.Source).SequenceEqual("---".AsSpan())) return;
        int start = header.Offset;
        while (header.TryRead(out SourceLine line)) {
            context.CancellationToken.ThrowIfCancellationRequested();
            ReadOnlySpan<char> marker = line.Text(body.Source).TrimEnd();
            if (!marker.SequenceEqual("---".AsSpan()) && !marker.SequenceEqual("...".AsSpan())) continue;
            if (line.Start - start > MaxAuthoringMetadataCharacters) {
                context.AddCell(NotebookCellKind.Raw, body.Source.Substring(opening.Start, header.Offset - opening.Start));
                context.Report("notebook-import-front-matter", "Oversized YAML front matter was retained in a raw cell.", Location(1),
                    OfficeConversionLossKind.Approximation);
                body = header;
                return;
            }
            string yaml = body.Source.Substring(start, line.Start - start);
            body = header;
            context.Document.SetMetadata("notebook_import_front_matter", NotebookImportContext.JsonString(yaml));
            ReadHeaderMetadata(yaml, ref language, context);
            return;
        }
        context.Report("notebook-import-front-matter", "The YAML front matter is missing its closing marker. The source was retained as Markdown.", Location(1));
    }

    private static void ReadHeaderMetadata(string yaml, ref string? language, NotebookImportContext context) {
        if (string.IsNullOrWhiteSpace(yaml)) return;
        try {
            using YamlDocument document = YamlDocument.Parse(yaml, YamlLoadOptions.UntrustedDefaults, context.CancellationToken);
            if (document.Root is not YamlMapping mapping) {
                context.Report("notebook-import-front-matter", "The YAML front matter was retained but is not a metadata mapping.", Location(1),
                    OfficeConversionLossKind.Approximation);
                return;
            }
            string? title = Scalar(mapping, "title");
            if (!string.IsNullOrWhiteSpace(title)) context.Document.SetMetadata("title", NotebookImportContext.JsonString(title!));
            if (mapping.TryGetValue("kernelspec", out YamlNode? spec) && spec is YamlMapping kernel) {
                language ??= Scalar(kernel, "language") ?? InferLanguage(Scalar(kernel, "name"));
            }
            language ??= InferLanguage(Scalar(mapping, "jupyter"));
        } catch (YamlException) {
            context.Report("notebook-import-front-matter", "The YAML front matter was retained as source because its metadata could not be read.", Location(1),
                OfficeConversionLossKind.Approximation);
        }
    }

    private static string? Scalar(YamlMapping mapping, string key) =>
        mapping.TryGetValue(key, out YamlNode? value) && value is YamlScalar scalar && scalar.TryGetValue(out string text) ? text : null;

    private static string? InferLanguage(string? kernelName) => kernelName?.ToLowerInvariant() switch {
        "python" or "python2" or "python3" or "ipython" => "python",
        "r" or "ir" => "r",
        "julia" => "julia",
        "ijavascript" => "javascript",
        _ => null
    };

    private static void SetDocumentLanguage(NotebookDocument notebook, string language) {
        using var payload = new MemoryStream();
        using (var writer = new Utf8JsonWriter(payload)) {
            writer.WriteStartObject();
            writer.WriteString("name", language);
            writer.WriteEndObject();
        }
        using var value = JsonDocument.Parse(payload.GetBuffer().AsMemory(0, (int)payload.Length));
        notebook.SetMetadata("language_info", value.RootElement);
    }
}
