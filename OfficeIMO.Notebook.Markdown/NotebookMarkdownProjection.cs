using System.Globalization;
using OfficeIMO.Markdown;

namespace OfficeIMO.Notebook.Markdown;

internal sealed partial class NotebookMarkdownProjection {
    private const int MaxCellTextCharacters = 8 * 1024 * 1024;
    private readonly NotebookDocument _notebook;
    private readonly NotebookMarkdownExportOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly List<IMarkdownBlock> _blocks = new();
    private readonly List<OfficeDiagnostic> _diagnostics;

    internal NotebookMarkdownProjection(NotebookDocument notebook, NotebookMarkdownExportOptions options, CancellationToken cancellationToken) {
        _notebook = notebook;
        _options = options;
        _cancellationToken = cancellationToken;
        _diagnostics = new List<OfficeDiagnostic>(notebook.Diagnostics);
    }

    internal IReadOnlyList<OfficeDiagnostic> Diagnostics => _diagnostics;

    internal MarkdownDocument Project() {
        if (!string.IsNullOrWhiteSpace(_options.Title)) {
            _blocks.Add(new HeadingBlock(1, new InlineSequence().Text(_options.Title!)));
        }
        if (!string.IsNullOrWhiteSpace(_options.Author)) {
            AddParagraph(_options.Author!);
        }

        for (int cellIndex = 0; cellIndex < _notebook.Cells.Count; cellIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            NotebookCell cell = _notebook.Cells[cellIndex];
            string location = $"cells[{cellIndex.ToString(CultureInfo.InvariantCulture)}]";
            switch (cell.Kind) {
                case NotebookCellKind.Markdown:
                    AddMarkdown(ReadSource(cell, location), new CellImages(this, cell, location), location);
                    break;
                case NotebookCellKind.Code:
                    AddCodeCell(cell, location);
                    break;
                case NotebookCellKind.Raw:
                    // Raw cells may contain target-specific markup. Keeping the source visible as literal text
                    // avoids accidentally treating it as recognized Markdown or executable content.
                    string raw = ReadSource(cell, location);
                    if (raw.Length > 0) _blocks.Add(new CodeBlock(string.Empty, raw));
                    break;
                default:
                    AddUnsupported("notebook-cell-type", location, "The cell type cannot be projected.");
                    AddParagraph("[Unsupported notebook cell]");
                    break;
            }
        }
        _cancellationToken.ThrowIfCancellationRequested();
        MarkdownDocument document = MarkdownDocument.Create();
        document.AddRange(_blocks);
        return document;
    }

    private void AddCodeCell(NotebookCell cell, string location) {
        if (_options.IncludeCode) {
            if (_options.IncludeExecutionCounts && cell.ExecutionCount is int executionCount) {
                AddParagraph("In [" + executionCount.ToString(CultureInfo.InvariantCulture) + "]");
            }
            string source = ReadSource(cell, location);
            if (source.Length > 0) _blocks.Add(new CodeBlock(LanguageLabel(cell.Language ?? _notebook.Language), source));
        }
        if (!_options.IncludeOutputs) return;

        for (int outputIndex = 0; outputIndex < cell.Outputs.Count; outputIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            NotebookOutput output = cell.Outputs[outputIndex];
            string outputLocation = location + ".outputs[" + outputIndex.ToString(CultureInfo.InvariantCulture) + "]";
            AddOutput(output, outputLocation);
        }
    }

    private string ReadSource(NotebookCell cell, string location) {
        string text = cell.GetSource(MaxCellTextCharacters, out bool truncated);
        if (truncated) AddUnsupported("notebook-cell-text-limit", location, "Cell text was truncated to the conversion limit.");
        return text;
    }

    private static string LanguageLabel(string? language) {
        if (string.IsNullOrWhiteSpace(language) || language!.Length > 64) return string.Empty;
        foreach (char character in language) {
            if (!char.IsLetterOrDigit(character) && character is not ('-' or '_' or '+' or '#' or '.')) return string.Empty;
        }
        return language;
    }

    private void AddParagraph(string text) => _blocks.Add(new ParagraphBlock(new InlineSequence().Text(text)));

    private void AddMarkdown(string text, CellImages? images, string location) {
        if (text.Length == 0) return;
        MarkdownLoadOptions settings = MarkdownLoadOptions.UntrustedDefaults;
        settings.CaptureSyntaxTree = false;
        settings.FrontMatter = false;
        settings.HtmlBlocks = false;
        settings.InlineHtml = false;
        settings.GenericAttributes = false;
        settings.AllowDataUrls = false;
        settings.DisallowFileUrls = true;
        settings.Math = true;
        // Preserve attachment reference definitions long enough for the image-specific resolver.
        settings.AllowedUrlSchemes = new[] { "http", "https", "mailto", "attachment" };
        settings.ImageSourceResolver = images != null ? images.Resolve : _ => RejectImage(location);
        if (settings.MaxInputCharacters is int maximum && text.Length > maximum) {
            AddUnsupported("notebook-markdown-limit", location, "The Markdown cell exceeds the safe parsing limit.");
            AddParagraph("[Markdown cell exceeds the parsing limit]");
            return;
        }
        using MarkdownDocument parsed = MarkdownDocument.Parse(text, settings);
        _blocks.AddRange(parsed.Blocks);
    }

    private string? RejectImage(string location) {
        AddUnsupported("notebook-image-reference", location, "The image has no supported cell-local attachment. External image references are not resolved.");
        return null;
    }

    private void AddUnsupported(string code, string location, string message, OfficeConversionLossKind kind = OfficeConversionLossKind.Omission) =>
        _diagnostics.Add(new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Warning, message, location, "Jupyter", kind));
}
