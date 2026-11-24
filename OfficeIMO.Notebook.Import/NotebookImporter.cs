using System.Diagnostics;

namespace OfficeIMO.Notebook.Import;

/// <summary>Imports saved notebook sources and outputs into the editable OfficeIMO notebook model.</summary>
public static partial class NotebookImporter {
    private const int CopyBufferSize = 64 * 1024;

    /// <summary>Imports a file using the selected format. The caller owns the resulting notebook.</summary>
    public static OfficeConversionResult<NotebookDocument> Load(string path, NotebookImportFormat format,
        NotebookImportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        cancellationToken.ThrowIfCancellationRequested();
        using FileStream stream = File.OpenRead(path);
        return Load(stream, format, options, cancellationToken);
    }

    /// <summary>Imports from the stream's current position and leaves the stream open.</summary>
    public static OfficeConversionResult<NotebookDocument> Load(Stream stream, NotebookImportFormat format,
        NotebookImportOptions? options = null, CancellationToken cancellationToken = default) {
        NotebookImportOptions effective = (options ?? new NotebookImportOptions()).Snapshot(format);
        var elapsed = Stopwatch.StartNew();
        using var source = new NotebookInputStream(stream, effective.Limits.MaxInputBytes, cancellationToken);
        using MemoryStream input = CreateBuffer(stream);
        source.CopyTo(input, CopyBufferSize);
        return Import(input.GetBuffer().AsMemory(0, (int)input.Length), format, effective, elapsed, cancellationToken);
    }

    /// <summary>Imports from bytes without retaining the caller's storage.</summary>
    public static OfficeConversionResult<NotebookDocument> Load(ReadOnlyMemory<byte> input, NotebookImportFormat format,
        NotebookImportOptions? options = null, CancellationToken cancellationToken = default) =>
        Import(input, format, (options ?? new NotebookImportOptions()).Snapshot(format), Stopwatch.StartNew(), cancellationToken);

    /// <summary>Reads a file asynchronously and imports it using the selected format.</summary>
    public static async Task<OfficeConversionResult<NotebookDocument>> LoadAsync(string path, NotebookImportFormat format,
        NotebookImportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 4096, useAsync: true);
        return await LoadAsync(stream, format, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reads input asynchronously from the current stream position and leaves the stream open.</summary>
    public static async Task<OfficeConversionResult<NotebookDocument>> LoadAsync(Stream stream, NotebookImportFormat format,
        NotebookImportOptions? options = null, CancellationToken cancellationToken = default) {
        NotebookImportOptions effective = (options ?? new NotebookImportOptions()).Snapshot(format);
        var elapsed = Stopwatch.StartNew();
        using var source = new NotebookInputStream(stream, effective.Limits.MaxInputBytes, cancellationToken);
        using MemoryStream input = CreateBuffer(stream);
        await source.CopyToAsync(input, CopyBufferSize, cancellationToken).ConfigureAwait(false);
        return Import(input.GetBuffer().AsMemory(0, (int)input.Length), format, effective, elapsed, cancellationToken);
    }

    private static OfficeConversionResult<NotebookDocument> Import(ReadOnlyMemory<byte> input, NotebookImportFormat format,
        NotebookImportOptions options, Stopwatch elapsed, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (input.Length > options.Limits.MaxInputBytes) throw NotebookInputStream.InputLimit(options.Limits.MaxInputBytes);
        if (input.Length >= 3 && input.Span[0] == 0xEF && input.Span[1] == 0xBB && input.Span[2] == 0xBF) input = input.Slice(3);
        var context = new NotebookImportContext(options, Label(format), cancellationToken);
        try {
            switch (format) {
                case NotebookImportFormat.PolyglotDib: NotebookScriptImporter.ReadDib(input, context); break;
                case NotebookImportFormat.JupytextPercent: NotebookScriptImporter.ReadJupytext(input, context); break;
                case NotebookImportFormat.Quarto: NotebookMarkupImporter.ReadQuarto(input, context); break;
                case NotebookImportFormat.MySTMarkdown: NotebookMarkupImporter.ReadMyST(input, context); break;
                case NotebookImportFormat.RMarkdown: NotebookMarkupImporter.ReadRMarkdown(input, context); break;
                case NotebookImportFormat.RNotebookHtml: RNotebookHtmlImporter.Read(input, context); break;
                case NotebookImportFormat.Deepnote: NotebookDeepnoteImporter.Read(input, context); break;
                case NotebookImportFormat.Zeppelin: NotebookZeppelinImporter.Read(input, context); break;
                case NotebookImportFormat.PowerShellScript: NotebookPowerShellScript.Read(input, context); break;
                default: throw new ArgumentOutOfRangeException(nameof(format));
            }
            return context.Complete(elapsed.Elapsed);
        } catch {
            context.Document.Dispose();
            throw;
        }
    }

    private static MemoryStream CreateBuffer(Stream stream) {
        // Seekable inputs are prevalidated by NotebookInputStream. Limit the initial allocation because
        // a stream can advertise more data than it delivers.
        long length = stream.CanSeek ? stream.Length - stream.Position : 0L;
        return new MemoryStream((int)Math.Min(Math.Max(0L, length), 1024 * 1024L));
    }

    private static string Label(NotebookImportFormat format) => format switch {
        NotebookImportFormat.PolyglotDib => ".NET Polyglot",
        NotebookImportFormat.JupytextPercent => "Jupytext percent script",
        NotebookImportFormat.Quarto => "Quarto",
        NotebookImportFormat.MySTMarkdown => "MyST Markdown",
        NotebookImportFormat.RMarkdown => "R Markdown",
        NotebookImportFormat.RNotebookHtml => "R Notebook HTML",
        NotebookImportFormat.Deepnote => "Deepnote",
        NotebookImportFormat.Zeppelin => "Apache Zeppelin",
        NotebookImportFormat.PowerShellScript => "PowerShell script",
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };
}
