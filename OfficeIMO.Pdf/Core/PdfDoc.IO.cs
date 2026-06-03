namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    /// <summary>
    /// Analyzes this generated document against its configured compliance profile using the same layout path as PDF output.
    /// </summary>
    public PdfComplianceReadinessReport AssessCompliance() => AssessCompliance(_options.ComplianceProfile);

    /// <summary>
    /// Analyzes this generated document against a formal compliance profile using generated font-usage evidence from layout.
    /// </summary>
    /// <param name="profile">Compliance profile to assess without enabling formal profile generation.</param>
    public PdfComplianceReadinessReport AssessCompliance(PdfComplianceProfile profile) {
        PdfGeneratedDocumentComplianceEvidence evidence = PdfWriter.CollectGeneratedComplianceEvidence(_blocks, _options);
        return PdfComplianceAnalyzer.AssessDocument(profile, _options, evidence.StandardFonts, _title, evidence.Images, evidence.Drawings, evidence.Forms);
    }

    /// <summary>
    /// Renders the document into a PDF byte array in memory.
    /// </summary>
    public byte[] ToBytes() => PdfWriter.Write(this, _blocks, _options, _title, _author, _subject, _keywords);

    /// <summary>
    /// Writes the document to <paramref name="stream"/> at the stream's current position.
    /// </summary>
    /// <param name="stream">Writable destination stream.</param>
    /// <returns>This <see cref="PdfDoc"/> for chaining.</returns>
    public PdfDoc Save(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));

        var bytes = ToBytes();
        stream.Write(bytes, 0, bytes.Length);
        return this;
    }

    /// <summary>
    /// Saves the document to <paramref name="path"/>. Creates the directory if needed.
    /// </summary>
    /// <param name="path">Destination file path, e.g. "C:\\Docs\\Report.pdf".</param>
    /// <returns>This <see cref="PdfDoc"/> for chaining.</returns>
    public PdfDoc Save(string path) {
        string fullPath = ValidateOutputPath(path);
        EnsureOutputDirectory(fullPath);

        var bytes = ToBytes();
        File.WriteAllBytes(fullPath, bytes);
        return this;
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
