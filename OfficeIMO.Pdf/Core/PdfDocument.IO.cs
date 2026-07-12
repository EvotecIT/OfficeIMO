using OfficeIMO.Drawing.Internal;
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
    /// Combines generated-document compliance readiness for the configured profile with external validator evidence.
    /// </summary>
    public PdfComplianceProofReport AssessComplianceProof(IEnumerable<PdfExternalValidationResult>? externalValidations = null) =>
        AssessComplianceProof(_options.ComplianceProfile, externalValidations);

    /// <summary>
    /// Combines generated-document compliance readiness for a formal profile with external validator evidence.
    /// </summary>
    /// <param name="profile">Compliance profile to assess without enabling formal profile generation.</param>
    /// <param name="externalValidations">Optional external validator results to combine with OfficeIMO.Pdf readiness evidence.</param>
    public PdfComplianceProofReport AssessComplianceProof(PdfComplianceProfile profile, IEnumerable<PdfExternalValidationResult>? externalValidations = null) {
        PdfComplianceReadinessReport readiness = AssessCompliance(profile);
        return PdfComplianceAnalyzer.AssessProof(readiness, externalValidations);
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
    /// Writes the complete document to <paramref name="stream"/>. Seekable streams are overwritten and rewound.
    /// </summary>
    /// <param name="stream">Writable destination stream.</param>
    /// <returns>This <see cref="PdfDocument"/> for chaining.</returns>
    public PdfDocument Save(Stream stream) {
        var bytes = ToBytes();
        OfficeStreamWriter.WriteAllBytes(stream, bytes);
        return this;
    }

    /// <summary>
    /// Attempts to write the document to <paramref name="stream"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(Stream stream) {
        try {
            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfSaveResult.Failed(outputPath: null, preflightException!);
            }

            var bytes = RenderBytesCore();
            OfficeStreamWriter.WriteAllBytes(stream, bytes);
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
        OfficeFileCommit.WriteAllBytes(fullPath, bytes);
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
            OfficeFileCommit.WriteAllBytes(fullPath, bytes);
            return PdfSaveResult.Success(fullPath, bytes.LongLength);
        } catch (Exception ex) {
            return PdfSaveResult.Failed(fullPath ?? path, ex);
        }
    }

    /// <summary>
    /// Asynchronously writes the complete document to <paramref name="stream"/>. Seekable streams are overwritten and rewound.
    /// </summary>
    public async System.Threading.Tasks.Task SaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();

        var bytes = ToBytes();
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Attempts to asynchronously write the document to <paramref name="stream"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        try {
            cancellationToken.ThrowIfCancellationRequested();

            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfSaveResult.Failed(outputPath: null, preflightException!);
            }

            var bytes = RenderBytesCore();
            await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
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
        await OfficeFileCommit.WriteAllBytesAsync(fullPath, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
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
            await OfficeFileCommit.WriteAllBytesAsync(fullPath, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
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
