namespace OfficeIMO.Pdf;

/// <summary>Controls fail-closed creation of normalized PDF repair artifacts.</summary>
public sealed class PdfRepairArtifactOptions {
    /// <summary>Cancellation observed between parsing, planning, rewriting, and verification stages.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    /// <summary>Rejects a clean source that has no recovered structural defects.</summary>
    public bool RequireRecoveredDefects { get; set; } = true;

    /// <summary>Rejects sources containing ambiguous defects that were detected but not recovered.</summary>
    public bool RejectDetectedOnlyDefects { get; set; } = true;

    /// <summary>Optional maximum byte count for the rewritten artifact.</summary>
    public long? MaximumOutputBytes { get; set; }
}

/// <summary>Normalized repaired PDF with before/after diagnostics and preservation evidence.</summary>
public sealed class PdfRepairArtifactResult {
    private readonly byte[] _pdf;
    private readonly PdfLoadOptions _readOptions;

    internal PdfRepairArtifactResult(
        byte[] pdf,
        long sourceSizeBytes,
        PdfRepairReport sourceRepairReport,
        PdfRepairReport strictOutputRepairReport,
        PdfRewritePreservationReport preservation,
        PdfLoadOptions readOptions) {
        _pdf = (byte[])pdf.Clone();
        SourceSizeBytes = sourceSizeBytes;
        SourceRepairReport = sourceRepairReport;
        StrictOutputRepairReport = strictOutputRepairReport;
        Preservation = preservation;
        _readOptions = readOptions;
    }

    /// <summary>Original source size in bytes.</summary>
    public long SourceSizeBytes { get; }
    /// <summary>Normalized artifact size in bytes.</summary>
    public long OutputSizeBytes => _pdf.LongLength;
    /// <summary>Every recovered and detected-only defect reported while opening the source.</summary>
    public PdfRepairReport SourceRepairReport { get; }
    /// <summary>Repair report from reopening the artifact under strict parsing. It must be empty.</summary>
    public PdfRepairReport StrictOutputRepairReport { get; }
    /// <summary>User-visible and structural preservation comparison.</summary>
    public PdfRewritePreservationReport Preservation { get; }
    /// <summary>True when strict reopening and preservation checks both succeeded.</summary>
    public bool IsVerified => !StrictOutputRepairReport.HasRepairs && Preservation.IsPreserved;
    /// <summary>Returns an independent copy of the normalized artifact.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the normalized artifact through the public document API.</summary>
    public PdfDocument ToDocument(PdfLoadOptions? readOptions = null) => PdfDocument.Load(_pdf, readOptions ?? _readOptions);
}

/// <summary>Creates bounded repair artifacts only for defects the parser recovered explicitly.</summary>
public static class PdfRepairArtifact {
    /// <summary>Normalizes recovered source defects, then proves strict reopen and structural preservation.</summary>
    public static PdfRepairArtifactResult Create(
        byte[] pdf,
        PdfRepairArtifactOptions? options = null,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfRepairArtifactOptions effective = options ?? new PdfRepairArtifactOptions();
        System.Threading.CancellationToken cancellationToken = effective.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        if (effective.MaximumOutputBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(options), "Maximum repair-artifact bytes must be positive.");

        PdfLoadOptions lenientOptions = CreateReadOptions(readOptions, PdfParsingMode.Lenient);
        PdfReadDocument source = PdfReadDocument.Open(pdf, lenientOptions);
        cancellationToken.ThrowIfCancellationRequested();
        PdfRepairReport sourceRepairs = source.RepairReport;
        if (effective.RequireRecoveredDefects && sourceRepairs.RepairCount == 0) {
            throw new InvalidOperationException("The source PDF has no recovered structural defects to persist.");
        }
        if (effective.RejectDetectedOnlyDefects && sourceRepairs.DetectionOnlyCount > 0) {
            throw new InvalidOperationException("The source PDF contains detected-only structural defects and cannot be repaired automatically.");
        }

        // Optimize is the existing full-rewrite authorization for normalized object graphs. The
        // repair artifact does not apply optimizer transforms; it only serializes the recovered graph.
        _ = PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.Optimize, lenientOptions);
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Security.HasEncryption) throw new NotSupportedException("Repair artifacts do not silently remove or replace source encryption.");
        if (source.Security.HasSignatures || source.Security.HasDocMDPPermissions || source.Security.HasUsageRights) {
            throw new NotSupportedException("Repair artifacts cannot preserve signatures, certification permissions, or usage rights through a full rewrite.");
        }

        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(
            pdf,
            lenientOptions,
            outputEncryption: null,
            mutateObjectGraph: null,
            maximumOutputBytes: effective.MaximumOutputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions strictOptions = CreateReadOptions(
            PdfLoadOptions.ForGeneratedOutput(readOptions, pdf, output),
            PdfParsingMode.Strict);
        PdfReadDocument strictOutput = PdfReadDocument.Open(output, strictOptions);
        cancellationToken.ThrowIfCancellationRequested();
        if (strictOutput.RepairReport.HasRepairs) throw new InvalidOperationException("The repaired artifact still requires parser recovery.");

        PdfRewritePreservationReport preservation = PdfRewritePreservation.Assess(
            pdf,
            output,
            options: null,
            originalReadOptions: lenientOptions,
            rewrittenReadOptions: strictOptions);
        cancellationToken.ThrowIfCancellationRequested();
        preservation.ThrowIfFailed();
        return new PdfRepairArtifactResult(
            output,
            pdf.LongLength,
            sourceRepairs,
            strictOutput.RepairReport,
            preservation,
            strictOptions);
    }

    private static PdfLoadOptions CreateReadOptions(PdfLoadOptions? readOptions, PdfParsingMode parsingMode) {
        PdfLoadOptions effective = PdfLoadOptions.Resolve(readOptions);
        return new PdfLoadOptions {
            ParsingMode = parsingMode,
            Limits = effective.Limits,
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }
}
