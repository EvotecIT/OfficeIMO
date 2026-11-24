using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Notebook.Markdown;

internal sealed partial class NotebookMarkdownProjection {
    private const int MaxHtmlOutputCharacters = 2 * 1024 * 1024;
    private static readonly string[] PreferredOutputTypes = { "text/html", "text/latex", "text/markdown", "text/plain" };

    private void AddOutput(NotebookOutput output, string location) {
        switch (output.Kind) {
            case NotebookOutputKind.Stream:
                string text = output.GetText(MaxCellTextCharacters, out bool truncated);
                if (truncated) AddUnsupported("notebook-output-text-limit", location, "Stream output text was truncated to the conversion limit.");
                if (text.Length > 0) {
                    AddLiteralOutput(text, output.Name == "stderr" ? "stderr" : "stdout");
                }
                break;
            case NotebookOutputKind.Error:
                string error = string.Join(": ", new[] { output.ErrorName, output.ErrorValue }.Where(part => !string.IsNullOrEmpty(part)));
                if (output.Traceback.Count > 0) {
                    if (error.Length > 0) error += "\n";
                    error += string.Join("\n", output.Traceback);
                }
                AddLiteralOutput(error, "Error");
                break;
            case NotebookOutputKind.ExecuteResult:
            case NotebookOutputKind.DisplayData:
                if (_options.IncludeExecutionCounts && output.ExecutionCount is int executionCount) {
                    AddParagraph("Out [" + executionCount.ToString(CultureInfo.InvariantCulture) + "]");
                }
                if (_options.OutputRenderer?.Invoke(output, location) is IMarkdownBlock rendered) {
                    _blocks.Add(rendered);
                    break;
                }
                AddDisplay(output.Data, location);
                break;
            default:
                AddUnsupported("notebook-output-type", location, "The saved output type cannot be projected.");
                AddParagraph("[Unsupported notebook output]");
                break;
        }
    }

    private void AddLiteralOutput(string value, string label) {
        // Jupyter traceback and stream strings can include ANSI SGR sequences. Strip only their
        // formatting, keeping visible data and line boundaries for PDF, Markdown, and Word.
        _blocks.Add(new CodeBlock(label, StripSgr(value)));
    }

    private static string StripSgr(string value) {
        int escape = value.IndexOf('\u001b');
        if (escape < 0) return value;
        var output = new StringBuilder(value.Length);
        output.Append(value, 0, escape);
        for (int index = escape; index < value.Length; index++) {
            if (value[index] == '\u001b' && index + 1 < value.Length && value[index + 1] == '[') {
                int end = index + 2;
                while (end < value.Length && (char.IsDigit(value[end]) || value[end] is ';' or ':')) end++;
                if (end < value.Length && value[end] == 'm') {
                    index = end;
                    continue;
                }
            }
            if (value[index] != '\u001b') output.Append(value[index]);
        }
        return output.ToString();
    }

    private void AddDisplay(IReadOnlyDictionary<string, JsonElement> data, string location) {
        if (data.Count == 0) return;
        bool usedImage = TryAddSavedDisplay(data, PreferredImageTypes, location);
        if (!usedImage && TryAddChart(data, location)) return;
        if (usedImage || TryAddSavedDisplay(data, PreferredOutputTypes, location)) {
            if (data.Keys.Any(IsInteractive)) {
                AddUnsupported("notebook-output-static-fallback", location, "A stored static representation was used for interactive output.", OfficeConversionLossKind.Approximation);
            }
            return;
        }
        AddUnsupported("notebook-output-unsupported", location, "The saved output has no supported static representation.");
        AddParagraph("[Saved notebook output unavailable]");
    }

    private bool TryAddSavedDisplay(IReadOnlyDictionary<string, JsonElement> data, string[] preferredTypes, string location) {
        foreach (string mime in preferredTypes) {
            if (!data.TryGetValue(mime, out JsonElement item)) continue;
            _cancellationToken.ThrowIfCancellationRequested();
            if (TryAddDisplay(mime, item, location)) return true;
            AddUnsupported("notebook-output-representation", location, "The saved " + mime + " representation is unsupported, invalid, or exceeds the conversion limit.");
        }
        return false;
    }

    private bool TryAddDisplay(string mime, JsonElement item, string location) {
        if (mime.StartsWith("image/", StringComparison.Ordinal)) {
            if (!TryGetImageUri(mime, item, out string imageUri)) return false;
            _blocks.Add(new ImageBlock(imageUri, "Notebook output image", null));
            return true;
        }
        int limit = mime == "text/html" ? MaxHtmlOutputCharacters : MaxCellTextCharacters;
        if (!TryReadMultiline(item, limit, out string value)) return false;
        switch (mime) {
            case "text/html":
                return TryAddHtmlTable(value, location);
            case "text/latex":
                _blocks.Add(new MathBlock(TrimMathDelimiters(value)));
                break;
            case "text/markdown":
                AddMarkdown(value, images: null, location);
                break;
            case "text/plain":
                AddLiteralOutput(value, string.Empty);
                break;
        }
        return true;
    }

    private bool TryAddHtmlTable(string value, string location) {
        // Render the bounded semantic table subset through the shared inert HTML importer.
        // An interactive graph usually provides HTML with no actual data table. Prefer its other MIME representations.
        if (value.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0) return false;
        HtmlConversionDocumentOptions settings = HtmlConversionDocumentOptions.UntrustedDefaults;
        settings.IncludeNormalizedHtml = false;
        settings.Limits.MaxInputCharacters = MaxHtmlOutputCharacters;
        settings.Limits.MaxHtmlNodes = 30_000;
        settings.Limits.MaxHtmlDepth = 64;
        HtmlUrlPolicy resourcePolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();
        resourcePolicy.ResolvedUrlTransform = _ => null;
        settings.ResourceUrlPolicy = resourcePolicy;
        try {
            HtmlConversionDocument html = HtmlConversionDocument.Parse(value, settings, _cancellationToken);
            using MarkdownDocument converted = html.ToMarkdown(new HtmlMarkdownExportOptions {
                RemoveScriptsAndStyles = true,
                PreserveUnsupportedBlocks = false,
                PreserveUnsupportedInlineHtml = false,
                UnknownBlockHandling = MarkdownHtmlUnknownTagHandling.Bypass,
                UnknownInlineHandling = MarkdownHtmlUnknownTagHandling.Bypass,
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                ResourceUrlPolicy = resourcePolicy,
                MaxInputCharacters = MaxHtmlOutputCharacters,
                MaxTableExpandedColumns = 256
            }, out OfficeConversionReport report, _cancellationToken);
            if (!converted.HasDescendantOfType<TableBlock>()) return false;
            foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
                _diagnostics.Add(new OfficeDiagnostic(diagnostic.Code, diagnostic.Severity, diagnostic.Message,
                    location, diagnostic.Format, diagnostic.LossKind));
            }
            _blocks.AddRange(converted.Blocks);
            return true;
        } catch (HtmlDomLimitException) {
            return false;
        }
    }

    private static bool IsInteractive(string mime) => mime is "application/javascript" or "text/javascript"
        || Array.IndexOf(VegaLiteMimes, mime) >= 0
        || mime != "application/vnd.dataresource+json" && mime.StartsWith("application/vnd.", StringComparison.Ordinal)
            && mime.EndsWith("+json", StringComparison.Ordinal);

    private static string TrimMathDelimiters(string value) {
        string trimmed = value.Trim();
        if (trimmed.Length >= 4 && (trimmed.StartsWith("$$", StringComparison.Ordinal) && trimmed.EndsWith("$$", StringComparison.Ordinal)
            || trimmed.StartsWith("\\[", StringComparison.Ordinal) && trimmed.EndsWith("\\]", StringComparison.Ordinal)
            || trimmed.StartsWith("\\(", StringComparison.Ordinal) && trimmed.EndsWith("\\)", StringComparison.Ordinal))) {
            return trimmed.Substring(2, trimmed.Length - 4);
        }
        if (trimmed.Length >= 2 && trimmed[0] == '$' && trimmed[trimmed.Length - 1] == '$') return trimmed.Substring(1, trimmed.Length - 2);
        return trimmed;
    }

    private static bool TryReadMultiline(JsonElement value, int maxCharacters, out string text) {
        text = string.Empty;
        if (value.ValueKind != JsonValueKind.String && (value.ValueKind != JsonValueKind.Array
            || value.EnumerateArray().Any(static fragment => fragment.ValueKind != JsonValueKind.String))) return false;
        text = NotebookJson.ReadText(value, maxCharacters, out bool truncated);
        return !truncated;
    }
}
