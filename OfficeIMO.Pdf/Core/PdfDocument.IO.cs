using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Analyzes this generated document against its configured compliance profile using the same layout path as PDF output.
    /// </summary>
    public PdfComplianceReadinessReport AssessCompliance() => AssessCompliance(_options.ComplianceProfile);

    /// <summary>
    /// Analyzes this document against a formal compliance profile.
    /// Generated documents use layout evidence; opened documents use artifact readback evidence.
    /// </summary>
    /// <param name="profile">Compliance profile to assess without enabling formal profile generation.</param>
    public PdfComplianceReadinessReport AssessCompliance(PdfComplianceProfile profile) {
        if (_source is not null) {
            var snapshot = GetReadSnapshot();
            PdfDocumentInfo info = PdfInspector.Inspect(snapshot.Bytes, snapshot.Document);
            return PdfComplianceAnalyzer.AssessReadback(profile, snapshot.Document, info);
        }

        PdfGeneratedDocumentComplianceEvidence evidence = PdfWriter.CollectGeneratedComplianceEvidence(this, _blocks, _options);
        return PdfComplianceAnalyzer.AssessDocument(profile, _options, evidence.StandardFonts, evidence.FontUsages, _title, evidence.Images, evidence.Drawings, evidence.Forms);
    }

    /// <summary>
    /// Combines generated-document compliance readiness for the configured profile with external validator evidence.
    /// </summary>
    public PdfComplianceProofReport AssessComplianceProof(IEnumerable<PdfExternalValidationResult>? externalValidations = null) =>
        AssessComplianceProof(_options.ComplianceProfile, externalValidations);

    /// <summary>
    /// Atomically renders or snapshots this document with readiness evidence for its configured compliance profile.
    /// </summary>
    public PdfComplianceArtifact CreateComplianceArtifact() =>
        CreateComplianceArtifact(_options.ComplianceProfile);

    /// <summary>
    /// Atomically renders or snapshots this document with readiness evidence for <paramref name="profile"/>.
    /// Use the returned artifact's bytes for external validation, then call its
    /// <see cref="PdfComplianceArtifact.AssessProof"/> method with those validator results.
    /// </summary>
    public PdfComplianceArtifact CreateComplianceArtifact(PdfComplianceProfile profile) {
        Guard.ComplianceProfile(profile, nameof(profile));
        if (_source is not null) {
            PdfComplianceReadinessReport openedReadiness = AssessCompliance(profile);
            return new PdfComplianceArtifact(_source.CopyBytes(), openedReadiness, ReadOptions);
        }

        ThrowIfTextEncodingPreflightFails();
        (byte[] bytes, PdfGeneratedDocumentComplianceEvidence evidence) = PdfWriter.WriteComplianceArtifact(
            this,
            _blocks,
            _options,
            _title,
            _author,
            _subject,
            _keywords);
        PdfComplianceReadinessReport generatedReadiness = PdfComplianceAnalyzer.AssessDocument(
            profile,
            _options,
            evidence.StandardFonts,
            evidence.FontUsages,
            _title,
            evidence.Images,
            evidence.Drawings,
            evidence.Forms);
        PdfStandardEncryptionOptions? encryption = _options.EncryptionSnapshot;
        PdfReadOptions? readOptions = encryption == null
            ? null
            : new PdfReadOptions { Password = encryption.UserPassword };
        return new PdfComplianceArtifact(bytes, generatedReadiness, readOptions);
    }

    /// <summary>
    /// Combines generated-document compliance readiness for a formal profile with external validator evidence.
    /// </summary>
    /// <param name="profile">Compliance profile to assess without enabling formal profile generation.</param>
    /// <param name="externalValidations">Optional external validator results to combine with OfficeIMO.Pdf readiness evidence.</param>
    public PdfComplianceProofReport AssessComplianceProof(PdfComplianceProfile profile, IEnumerable<PdfExternalValidationResult>? externalValidations = null) {
        PdfComplianceReadinessReport readiness = AssessCompliance(profile);
        return _source is null
            ? PdfComplianceAnalyzer.AssessProof(readiness, externalValidations)
            : PdfComplianceAnalyzer.AssessProof(readiness, GetBytesForOperation(), externalValidations);
    }

    /// <summary>
    /// Renders the document into a PDF byte array in memory.
    /// </summary>
    public byte[] ToBytes() {
        if (_source is not null) {
            return _source.CopyBytes();
        }

        ThrowIfTextEncodingPreflightFails();
        return RenderBytesCore();
    }

    /// <summary>Renders the document into a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream() => new MemoryStream(ToBytes());

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
    public void Save(Stream stream) {
        ThrowIfTextEncodingPreflightFails();
        RenderToStreamCore(stream);
    }

    /// <summary>
    /// Attempts to write the document to <paramref name="stream"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(Stream stream) {
        try {
            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                return PdfSaveResult.Failed(outputPath: null, preflightException!);
            }

            long bytesWritten = RenderToStreamCore(stream);
            return PdfSaveResult.Success(outputPath: null, bytesWritten);
        } catch (Exception ex) {
            return PdfSaveResult.Failed(outputPath: null, ex);
        }
    }

    /// <summary>
    /// Saves the document to <paramref name="path"/>. Creates the directory if needed.
    /// </summary>
    /// <param name="path">Destination file path, e.g. "C:\\Docs\\Report.pdf".</param>
    public void Save(string path) {
        string fullPath = ValidateOutputPath(path);
        EnsureOutputDirectory(fullPath);

        ThrowIfTextEncodingPreflightFails();
        OfficeFileCommit.Write(fullPath, stream => WritePdfCore(stream));
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

            long bytesWritten = 0L;
            OfficeFileCommit.Write(fullPath, stream => bytesWritten = WritePdfCore(stream));
            return PdfSaveResult.Success(fullPath, bytesWritten);
        } catch (Exception ex) {
            return PdfSaveResult.Failed(fullPath ?? path, ex);
        }
    }

    /// <summary>
    /// Asynchronously writes the complete document to <paramref name="stream"/>. Seekable streams are overwritten and rewound.
    /// </summary>
    public async System.Threading.Tasks.Task SaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ThrowIfTextEncodingPreflightFails();
        await RenderToStreamCoreAsync(stream, cancellationToken).ConfigureAwait(false);
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

            long bytesWritten = await RenderToStreamCoreAsync(stream, cancellationToken).ConfigureAwait(false);
            return PdfSaveResult.Success(outputPath: null, bytesWritten);
        } catch (System.OperationCanceledException) {
            throw;
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

        ThrowIfTextEncodingPreflightFails();
        await OfficeFileCommit.WriteAsync(fullPath, stream => WritePdfCore(stream), cancellationToken: cancellationToken).ConfigureAwait(false);
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

            long bytesWritten = 0L;
            await OfficeFileCommit.WriteAsync(fullPath, stream => bytesWritten = WritePdfCore(stream), cancellationToken: cancellationToken).ConfigureAwait(false);
            return PdfSaveResult.Success(fullPath, bytesWritten);
        } catch (System.OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            return PdfSaveResult.Failed(fullPath ?? path, ex);
        }
    }

    private byte[] RenderBytesCore() {
        if (_source is not null) {
            return _source.CopyBytes();
        }

        return PdfWriter.Write(this, _blocks, _options, _title, _author, _subject, _keywords);
    }

    private long RenderToStreamCore(Stream stream) {
        long bytesWritten = 0L;
        OfficeStreamWriter.Write(stream, destination => bytesWritten = WritePdfCore(destination));
        return bytesWritten;
    }

    private async System.Threading.Tasks.Task<long> RenderToStreamCoreAsync(Stream stream, System.Threading.CancellationToken cancellationToken) {
        long bytesWritten = 0L;
        await OfficeStreamWriter.WriteAsync(
            stream,
            destination => bytesWritten = WritePdfCore(destination),
            cancellationToken).ConfigureAwait(false);
        return bytesWritten;
    }

    private long WritePdfCore(Stream stream) {
        if (_source is not null) {
            stream.Write(_source.Bytes, 0, _source.Bytes.Length);
            return _source.Bytes.LongLength;
        }

        return PdfWriter.Write(stream, this, _blocks, _options, _title, _author, _subject, _keywords);
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
