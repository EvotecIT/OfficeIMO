namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Analyzes this generated document against its configured compliance profile using the same layout path as PDF output.
    /// </summary>
    public PdfComplianceReadinessReport AssessCompliance() => AssessCompliance(_options.ComplianceProfile);

    /// <summary>
    /// Analyzes this generated document against a formal compliance profile using generated font-usage evidence from layout.
    /// </summary>
    /// <param name="profile">Compliance profile to assess without enabling formal profile generation.</param>
    public PdfComplianceReadinessReport AssessCompliance(PdfComplianceProfile profile) {
        EnsureGeneratedDocument();
        PdfGeneratedDocumentComplianceEvidence evidence = PdfWriter.CollectGeneratedComplianceEvidence(_blocks, _options);
        return PdfComplianceAnalyzer.AssessDocument(profile, _options, evidence.StandardFonts, evidence.FontUsages, _title, evidence.Images, evidence.Drawings, evidence.Forms);
    }

    /// <summary>
    /// Renders the document into a PDF byte array in memory.
    /// </summary>
    public byte[] ToBytes() {
        if (_loadedPdf is not null) {
            return (byte[])_loadedPdf.Clone();
        }

        ThrowIfTextEncodingPreflightFails();
        return RenderBytesCore();
    }

    /// <summary>
    /// Attempts to render the document into a PDF byte array and returns diagnostics instead of throwing.
    /// </summary>
    public PdfBytesResult TryToBytes() {
        try {
            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfBytesResult.Failed(preflightException!);
            }

            return PdfBytesResult.Success(RenderBytesCore());
        } catch (Exception ex) {
            return PdfBytesResult.Failed(ex);
        }
    }

    /// <summary>
    /// Writes the document to <paramref name="stream"/> at the stream's current position.
    /// </summary>
    /// <param name="stream">Writable destination stream.</param>
    /// <returns>This <see cref="PdfDocument"/> for chaining.</returns>
    public PdfDocument Save(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));

        var bytes = ToBytes();
        stream.Write(bytes, 0, bytes.Length);
        return this;
    }

    /// <summary>
    /// Attempts to write the document to <paramref name="stream"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(Stream stream) {
        try {
            Guard.NotNull(stream, nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));

            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfSaveResult.Failed(outputPath: null, preflightException!);
            }

            var bytes = RenderBytesCore();
            stream.Write(bytes, 0, bytes.Length);
            return PdfSaveResult.Success(outputPath: null, bytes.LongLength);
        } catch (Exception ex) {
            return PdfSaveResult.Failed(outputPath: null, ex);
        }
    }

    /// <summary>
    /// Saves the document to <paramref name="path"/>. Creates the directory if needed.
    /// </summary>
    /// <param name="path">Destination file path, e.g. "C:\\Docs\\Report.pdf".</param>
    /// <returns>This <see cref="PdfDocument"/> for chaining.</returns>
    public PdfDocument Save(string path) {
        string fullPath = ValidateOutputPath(path);
        EnsureOutputDirectory(fullPath);

        var bytes = ToBytes();
        File.WriteAllBytes(fullPath, bytes);
        return this;
    }

    /// <summary>
    /// Attempts to save the document to <paramref name="path"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(string path) {
        string? fullPath = null;
        try {
            fullPath = ValidateOutputPath(path);
            EnsureOutputDirectory(fullPath);

            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfSaveResult.Failed(fullPath, preflightException!);
            }

            var bytes = RenderBytesCore();
            File.WriteAllBytes(fullPath, bytes);
            return PdfSaveResult.Success(fullPath, bytes.LongLength);
        } catch (Exception ex) {
            return PdfSaveResult.Failed(fullPath ?? path, ex);
        }
    }

    /// <summary>
    /// Asynchronously writes the document to <paramref name="stream"/> at the stream's current position.
    /// </summary>
    public async System.Threading.Tasks.Task SaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();

        var bytes = ToBytes();
#if NET8_0_OR_GREATER
        await stream.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }

    /// <summary>
    /// Attempts to asynchronously write the document to <paramref name="stream"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        try {
            Guard.NotNull(stream, nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
            cancellationToken.ThrowIfCancellationRequested();

            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfSaveResult.Failed(outputPath: null, preflightException!);
            }

            var bytes = RenderBytesCore();
#if NET8_0_OR_GREATER
            await stream.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
            await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
            return PdfSaveResult.Success(outputPath: null, bytes.LongLength);
        } catch (Exception ex) {
            return PdfSaveResult.Failed(outputPath: null, ex);
        }
    }

    /// <summary>
    /// Asynchronously saves the document to <paramref name="path"/>.
    /// </summary>
    public async System.Threading.Tasks.Task SaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        string fullPath = ValidateOutputPath(path);
        cancellationToken.ThrowIfCancellationRequested();
        EnsureOutputDirectory(fullPath);

        var bytes = ToBytes();
        using var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
#if NET8_0_OR_GREATER
        await fs.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
        await fs.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }

    /// <summary>
    /// Attempts to asynchronously save the document to <paramref name="path"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        string? fullPath = null;
        try {
            fullPath = ValidateOutputPath(path);
            cancellationToken.ThrowIfCancellationRequested();
            EnsureOutputDirectory(fullPath);

            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfSaveResult.Failed(fullPath ?? path, preflightException!);
            }

            var bytes = RenderBytesCore();
            using var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
#if NET8_0_OR_GREATER
            await fs.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
            await fs.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
            return PdfSaveResult.Success(fullPath, bytes.LongLength);
        } catch (Exception ex) {
            return PdfSaveResult.Failed(fullPath ?? path, ex);
        }
    }

    private byte[] RenderBytesCore() {
        if (_loadedPdf is not null) {
            return (byte[])_loadedPdf.Clone();
        }

        return PdfWriter.Write(this, _blocks, _options, _title, _author, _subject, _keywords);
    }

    private void ThrowIfTextEncodingPreflightFails() {
        if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
            throw preflightException!;
        }
    }

    private static string ValidateOutputPath(string path) {
        Guard.NotNull(path, nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path cannot be empty or whitespace.", nameof(path));

        string fullPath;
        try { fullPath = System.IO.Path.GetFullPath(path); } catch (Exception ex) { throw new ArgumentException("Path is invalid.", nameof(path), ex); }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory)
            throw new ArgumentException("Path refers to a directory; a file path is required.", nameof(path));

        var fileName = System.IO.Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) throw new ArgumentException("Path must include a file name.", nameof(path));
        if (fileName.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0)
            throw new ArgumentException("Path contains invalid file name characters.", nameof(path));

        return fullPath;
    }

    private static void EnsureOutputDirectory(string fullPath) {
        var dir = System.IO.Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(dir)) dir = ".";
        Directory.CreateDirectory(dir);
    }
}
