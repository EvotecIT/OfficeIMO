using System.Diagnostics;
using OfficeIMO.Notebook.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Notebook.Word;

/// <summary>Converts notebook cells and saved output into editable Word content.</summary>
public static class NotebookWordConverterExtensions {
    /// <summary>Creates a detached Word document from saved notebook content.</summary>
    public static WordDocument ToWord(this NotebookDocument notebook, NotebookWordExportOptions? options = null,
        CancellationToken cancellationToken = default) => notebook.ToWord(options, out _, cancellationToken);

    /// <summary>Creates an editable Word document and returns projection and Word conversion findings.</summary>
    public static WordDocument ToWord(this NotebookDocument notebook, NotebookWordExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(notebook);
        options = (options ?? new NotebookWordExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        using var projection = notebook.ToMarkdown(new NotebookMarkdownExportOptions {
            IncludeCode = options.IncludeCode,
            IncludeOutputs = options.IncludeOutputs,
            IncludeExecutionCounts = options.IncludeExecutionCounts,
            Title = options.Title,
            Author = options.Author
        }, out OfficeConversionReport sourceReport, cancellationToken);
        WordDocument word = projection.ToWord(new MarkdownWordExportOptions {
            AllowLocalImages = false,
            AllowRemoteImages = false,
            AllowDataUriImages = true
        }, out OfficeConversionReport targetReport, cancellationToken);
        try {
            report = new OfficeConversionReport("Jupyter Notebook", "DOCX", itemsWritten: sourceReport.ItemsWritten,
                elapsed: timer.Elapsed, diagnostics: sourceReport.Diagnostics.Concat(targetReport.Diagnostics));
            if (options.FailOnLoss) report.ThrowIfLoss();
            cancellationToken.ThrowIfCancellationRequested();
            return word;
        } catch {
            word.Dispose();
            throw;
        }
    }

    /// <summary>Creates an editable Word document and returns the model and conversion report together.</summary>
    public static Task<OfficeConversionResult<WordDocument>> ToWordAsync(this NotebookDocument notebook,
        NotebookWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        NotebookWordExportOptions? snapshot = options?.Clone();
        return Task.Run(() => {
            WordDocument word = notebook.ToWord(snapshot, out OfficeConversionReport report, cancellationToken);
            return new OfficeConversionResult<WordDocument>(word, report);
        }, cancellationToken);
    }

    /// <summary>Saves an editable DOCX report to a path.</summary>
    public static OfficeConversionReport SaveAsWord(this NotebookDocument notebook, string path,
        NotebookWordExportOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.SaveAsWordAsync(path, options, cancellationToken).GetAwaiter().GetResult();

    /// <summary>Saves an editable DOCX report to a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsWord(this NotebookDocument notebook, Stream stream,
        NotebookWordExportOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.SaveAsWordAsync(stream, options, cancellationToken).GetAwaiter().GetResult();

    /// <summary>Saves an editable DOCX report to a path after projection and loss checks complete.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this NotebookDocument notebook, string path,
        NotebookWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("The notebook Word report requires a .docx destination.", nameof(path));
        }
        var converted = await notebook.ToWordAsync(options, cancellationToken).ConfigureAwait(false);
        using WordDocument word = converted.Value;
        await word.SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return new OfficeConversionReport(converted.Report.Source, converted.Report.Target, path,
            converted.Report.ItemsWritten, converted.Report.Elapsed, converted.Report.Diagnostics);
    }

    /// <summary>Saves an editable DOCX report to a caller-owned stream after loss checks complete.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this NotebookDocument notebook, Stream stream,
        NotebookWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        var converted = await notebook.ToWordAsync(options, cancellationToken).ConfigureAwait(false);
        using WordDocument word = converted.Value;
        await word.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return converted.Report;
    }
}
