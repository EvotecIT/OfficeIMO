using OfficeIMO.Markdown;
using OfficeIMO.Notebook.Markdown;

namespace OfficeIMO.Notebook.Html;

/// <summary>Controls the presentation of stored notebook content in offline HTML.</summary>
public sealed class NotebookHtmlExportOptions {
    private NotebookMarkdownExportOptions _projectionOptions = new NotebookMarkdownExportOptions();

    /// <summary>Notebook content selection, titles, and image limits.</summary>
    public NotebookMarkdownExportOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Produces an HTML fragment without a document envelope.</summary>
    public bool Fragment { get; set; }

    /// <summary>Built-in presentation style. Standalone documents embed the stylesheet.</summary>
    public MarkdownHtmlStyle Style { get; set; } = MarkdownHtmlStyle.Clean;

    /// <summary>Renders supported saved Jupyter widgets as offline browser controls.</summary>
    public bool IncludeInteractiveWidgets { get; set; }

    /// <summary>Rejects the result if a completed stage reports fidelity loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent snapshot of the conversion options.</summary>
    public NotebookHtmlExportOptions Clone() => new NotebookHtmlExportOptions {
        ProjectionOptions = ProjectionOptions.Clone(),
        Fragment = Fragment,
        Style = Style,
        IncludeInteractiveWidgets = IncludeInteractiveWidgets,
        FailOnLoss = FailOnLoss
    };
}
