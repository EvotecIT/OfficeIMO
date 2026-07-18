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
            : PdfComplianceAnalyzer.AssessProof(readiness, _source.Bytes, externalValidations, ReadOptions);
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
        var timer = System.Diagnostics.Stopwatch.StartNew();
        try {
            if (TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? preflightException)) {
                timer.Stop();
                PdfPipelineReport failedPipeline = AppendOutputStep("ToBytes", output: null, timer.Elapsed, preflightException);
                return PdfBytesResult.Failed(preflightException!, failedPipeline);
            }

            byte[] bytes = RenderBytesCore();
            timer.Stop();
            PdfArtifactSnapshot output = PdfArtifactSnapshot.Capture(bytes, ReadOptions);
            PdfPipelineReport pipeline = AppendOutputStep("ToBytes", output, timer.Elapsed);
            return PdfBytesResult.Success(bytes, pipeline);
        } catch (Exception ex) {
            timer.Stop();
            PdfPipelineReport pipeline = AppendOutputStep("ToBytes", output: null, timer.Elapsed, ex);
            return PdfBytesResult.Failed(ex, pipeline);
        }
    }

    /// <summary>
    /// Writes the complete document to <paramref name="stream"/>. Seekable streams are overwritten and rewound.
    /// </summary>
    /// <param name="stream">Writable destination stream.</param>
    public PdfSaveResult Save(Stream stream) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        ThrowIfTextEncodingPreflightFails();
        (long bytesWritten, PdfArtifactSnapshot output) = RenderToStreamWithEvidence(stream);
        timer.Stop();
        PdfPipelineReport pipeline = AppendOutputStep("Save", output, timer.Elapsed);
        return PdfSaveResult.Success(outputPath: null, bytesWritten, pipeline);
    }

    /// <summary>
    /// Attempts to write the document to <paramref name="stream"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(Stream stream) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        try {
            return Save(stream);
        } catch (Exception ex) {
            timer.Stop();
            PdfPipelineReport pipeline = AppendOutputStep("Save", output: null, timer.Elapsed, ex);
            return PdfSaveResult.Failed(outputPath: null, ex, pipeline);
        }
    }

    /// <summary>
    /// Saves the document to <paramref name="path"/>. Creates the directory if needed.
    /// </summary>
    /// <param name="path">Destination file path, e.g. "C:\\Docs\\Report.pdf".</param>
    public PdfSaveResult Save(string path) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        string fullPath = ValidateOutputPath(path);
        EnsureOutputDirectory(fullPath);

        ThrowIfTextEncodingPreflightFails();
        PdfArtifactSnapshot? output = null;
        long bytesWritten = 0L;
        OfficeFileCommit.Write(fullPath, stream => {
            using var hashingStream = new PdfPipelineHashingStream(stream);
            (bytesWritten, int? pageCount) = WritePdfCore(hashingStream);
            output = hashingStream.Complete(pageCount);
        });
        timer.Stop();
        PdfPipelineReport pipeline = AppendOutputStep("Save", output, timer.Elapsed);
        return PdfSaveResult.Success(fullPath, bytesWritten, pipeline);
    }

    /// <summary>
    /// Attempts to save the document to <paramref name="path"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(string path) {
        string? fullPath = null;
        var timer = System.Diagnostics.Stopwatch.StartNew();
        try {
            fullPath = ValidateOutputPath(path);
            return Save(fullPath);
        } catch (Exception ex) {
            timer.Stop();
            PdfPipelineReport pipeline = AppendOutputStep("Save", output: null, timer.Elapsed, ex);
            return PdfSaveResult.Failed(fullPath ?? path, ex, pipeline);
        }
    }

    /// <summary>
    /// Asynchronously writes the complete document to <paramref name="stream"/>. Seekable streams are overwritten and rewound.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> SaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        cancellationToken.ThrowIfCancellationRequested();
        ThrowIfTextEncodingPreflightFails();
        (long bytesWritten, PdfArtifactSnapshot output) = await RenderToStreamWithEvidenceAsync(stream, cancellationToken).ConfigureAwait(false);
        timer.Stop();
        PdfPipelineReport pipeline = AppendOutputStep("Save", output, timer.Elapsed);
        return PdfSaveResult.Success(outputPath: null, bytesWritten, pipeline);
    }

    /// <summary>
    /// Attempts to asynchronously write the document to <paramref name="stream"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        try {
            return await SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        } catch (System.OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            timer.Stop();
            PdfPipelineReport pipeline = AppendOutputStep("Save", output: null, timer.Elapsed, ex);
            return PdfSaveResult.Failed(outputPath: null, ex, pipeline);
        }
    }

    /// <summary>
    /// Asynchronously saves the document to <paramref name="path"/>.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> SaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        string fullPath = ValidateOutputPath(path);
        cancellationToken.ThrowIfCancellationRequested();
        EnsureOutputDirectory(fullPath);

        ThrowIfTextEncodingPreflightFails();
        PdfArtifactSnapshot? output = null;
        long bytesWritten = 0L;
        await OfficeFileCommit.WriteAsync(
            fullPath,
            stream => {
                using var hashingStream = new PdfPipelineHashingStream(stream);
                (bytesWritten, int? pageCount) = WritePdfCore(hashingStream);
                output = hashingStream.Complete(pageCount);
            },
            cancellationToken: cancellationToken).ConfigureAwait(false);
        timer.Stop();
        PdfPipelineReport pipeline = AppendOutputStep("Save", output, timer.Elapsed);
        return PdfSaveResult.Success(fullPath, bytesWritten, pipeline);
    }

    /// <summary>
    /// Attempts to asynchronously save the document to <paramref name="path"/> and returns output diagnostics instead of throwing.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        string? fullPath = null;
        var timer = System.Diagnostics.Stopwatch.StartNew();
        try {
            fullPath = ValidateOutputPath(path);
            return await SaveAsync(fullPath, cancellationToken).ConfigureAwait(false);
        } catch (System.OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            timer.Stop();
            PdfPipelineReport pipeline = AppendOutputStep("Save", output: null, timer.Elapsed, ex);
            return PdfSaveResult.Failed(fullPath ?? path, ex, pipeline);
        }
    }

    private byte[] RenderBytesCore() {
        if (_source is not null) {
            return _source.CopyBytes();
        }

        return PdfWriter.Write(this, _blocks, _options, _title, _author, _subject, _keywords);
    }

    private (long BytesWritten, int? PageCount) RenderToStreamCore(Stream stream) {
        (long BytesWritten, int? PageCount) output = default;
        OfficeStreamWriter.Write(stream, destination => output = WritePdfCore(destination));
        return output;
    }

    private async System.Threading.Tasks.Task<(long BytesWritten, int? PageCount)> RenderToStreamCoreAsync(
        Stream stream,
        System.Threading.CancellationToken cancellationToken) {
        (long BytesWritten, int? PageCount) output = default;
        await OfficeStreamWriter.WriteAsync(
            stream,
            destination => output = WritePdfCore(destination),
            cancellationToken).ConfigureAwait(false);
        return output;
    }

    private (long BytesWritten, PdfArtifactSnapshot Output) RenderToStreamWithEvidence(Stream stream) {
        using var hashingStream = new PdfPipelineHashingStream(stream);
        (long bytesWritten, int? pageCount) = RenderToStreamCore(hashingStream);
        PdfArtifactSnapshot output = hashingStream.Complete(pageCount);
        return (bytesWritten, output);
    }

    private async System.Threading.Tasks.Task<(long BytesWritten, PdfArtifactSnapshot Output)> RenderToStreamWithEvidenceAsync(
        Stream stream,
        System.Threading.CancellationToken cancellationToken) {
        using var hashingStream = new PdfPipelineHashingStream(stream);
        (long bytesWritten, int? pageCount) = await RenderToStreamCoreAsync(hashingStream, cancellationToken).ConfigureAwait(false);
        PdfArtifactSnapshot output = hashingStream.Complete(pageCount);
        return (bytesWritten, output);
    }

    private (long BytesWritten, int? PageCount) WritePdfCore(Stream stream) {
        if (_source is not null) {
            stream.Write(_source.Bytes, 0, _source.Bytes.Length);
            return (_source.Bytes.LongLength, _pipeline.Output?.PageCount);
        }

        long bytesWritten = PdfWriter.Write(
            stream,
            this,
            _blocks,
            _options,
            _title,
            _author,
            _subject,
            _keywords,
            out int pageCount);
        return (bytesWritten, pageCount);
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
