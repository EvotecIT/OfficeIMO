using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Notebook.Markdown;
using OfficeIMO.Pdf;

namespace OfficeIMO.Notebook.Pdf;

/// <summary>Controls presentation and layout when converting saved notebook cells to PDF.</summary>
public sealed class NotebookPdfExportOptions {
    private MarkdownPdfExportOptions _markdownOptions = new MarkdownPdfExportOptions {
        ResourcePolicy = PdfResourcePolicy.CreateLibraryFontDefault()
    };

    /// <summary>Includes code inputs. This setting is independent of saved outputs.</summary>
    public bool IncludeCode { get; set; } = true;

    /// <summary>Includes saved notebook outputs. Cells are never executed during conversion.</summary>
    public bool IncludeOutputs { get; set; } = true;

    /// <summary>Includes stored execution counts alongside the corresponding inputs and results.</summary>
    public bool IncludeExecutionCounts { get; set; } = true;

    /// <summary>Shows a title above the notebook and sets the PDF title.</summary>
    public string? Title { get; set; }

    /// <summary>Shows an author label above the notebook and sets the PDF author.</summary>
    public string? Author { get; set; }

    /// <summary>Rejects the artifact before publication if a completed conversion stage reports fidelity loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Shared PDF layout, pagination, fonts, and image limits. Notebook export always disables file and remote image resolution.</summary>
    public MarkdownPdfExportOptions MarkdownOptions {
        get => _markdownOptions;
        set => _markdownOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Creates an independent snapshot of the notebook export and layout settings.</summary>
    public NotebookPdfExportOptions Clone() {
        var copy = (NotebookPdfExportOptions)MemberwiseClone();
        copy._markdownOptions = _markdownOptions.Clone();
        return copy;
    }

    internal NotebookMarkdownExportOptions CreateProjectionOptions() => new NotebookMarkdownExportOptions {
        IncludeCode = IncludeCode,
        IncludeOutputs = IncludeOutputs,
        IncludeExecutionCounts = IncludeExecutionCounts,
        Title = Title,
        Author = Author,
        MaxImageBytes = MarkdownOptions.MaximumDataUriImageBytes
    };
}
