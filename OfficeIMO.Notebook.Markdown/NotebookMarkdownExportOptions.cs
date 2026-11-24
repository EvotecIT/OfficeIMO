namespace OfficeIMO.Notebook.Markdown;

/// <summary>Controls presentation of stored notebook cells and outputs in Markdown.</summary>
public sealed class NotebookMarkdownExportOptions {
    private int _maxImageBytes = 5 * 1024 * 1024;

    /// <summary>Includes code inputs. This setting is independent of saved outputs.</summary>
    public bool IncludeCode { get; set; } = true;

    /// <summary>Includes saved notebook outputs. Cells are never executed during conversion.</summary>
    public bool IncludeOutputs { get; set; } = true;

    /// <summary>Includes stored execution counts alongside the corresponding inputs and results.</summary>
    public bool IncludeExecutionCounts { get; set; } = true;

    /// <summary>Shows an explicit title above the notebook content.</summary>
    public string? Title { get; set; }

    /// <summary>Shows an explicit author label above the notebook content.</summary>
    public string? Author { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when conversion reports unrequested fidelity loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Maximum decoded size, in bytes, of each notebook image or attachment.</summary>
    public int MaxImageBytes {
        get => _maxImageBytes;
        set {
            if (value < 1) throw new ArgumentOutOfRangeException(nameof(value));
            _maxImageBytes = value;
        }
    }

    /// <summary>Creates an independent snapshot of these options.</summary>
    public NotebookMarkdownExportOptions Clone() => (NotebookMarkdownExportOptions)MemberwiseClone();
}
