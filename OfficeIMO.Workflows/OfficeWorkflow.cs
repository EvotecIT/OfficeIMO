namespace OfficeIMO.Workflows;

/// <summary>Creates fluent local document workflows over <see cref="OfficeWorkflowRunner"/>.</summary>
public static partial class OfficeWorkflow {






    /// <summary>Creates a PDF inspection workflow.</summary>
    public static OfficeWorkflowBuilder Inspect(string inputPath) =>
        Create(OfficeWorkflowOperation.Inspect, inputPath);

    /// <summary>Creates a PDF comparison workflow.</summary>
    public static OfficeWorkflowBuilder Compare(string inputPath, string comparisonPath) =>
        Create(OfficeWorkflowOperation.Compare, inputPath).Against(comparisonPath);

    /// <summary>Creates a lossless PDF optimization workflow.</summary>
    public static OfficeWorkflowBuilder Optimize(string inputPath) =>
        Create(OfficeWorkflowOperation.Optimize, inputPath);

    /// <summary>Creates a report-only PDF repair-planning workflow.</summary>
    public static OfficeWorkflowBuilder PlanRepair(string inputPath) =>
        Create(OfficeWorkflowOperation.RepairPlan, inputPath);

    /// <summary>Creates a verified PDF repair workflow.</summary>
    public static OfficeWorkflowBuilder Repair(string inputPath) =>
        Create(OfficeWorkflowOperation.Repair, inputPath);

    /// <summary>Creates a verified PDF sanitization workflow.</summary>
    public static OfficeWorkflowBuilder Sanitize(string inputPath) =>
        Create(OfficeWorkflowOperation.Sanitize, inputPath);

    /// <summary>Creates a PDF extraction workflow preserving the supplied page order and repeats.</summary>
    public static OfficeWorkflowBuilder ExtractPages(string inputPath, params int[] pageNumbers) =>
        new(new OfficeWorkflowRequest { Operation = OfficeWorkflowOperation.ExtractPages,
            InputPath = inputPath, PageNumbers = pageNumbers?.ToArray() ?? throw new ArgumentNullException(nameof(pageNumbers)) });

    /// <summary>Creates a password-protected PDF copy using captured output settings.</summary>
    public static OfficeWorkflowBuilder ProtectPdf(string inputPath, OfficeIMO.Pdf.PdfStandardEncryptionOptions encryption) =>
        new(new OfficeWorkflowRequest { Operation = OfficeWorkflowOperation.ProtectPdf, InputPath = inputPath,
            OutputEncryption = encryption?.Clone() ?? throw new ArgumentNullException(nameof(encryption)) });

    /// <summary>Creates an unencrypted PDF copy with the current owner password.</summary>
    public static OfficeWorkflowBuilder RemovePdfProtection(string inputPath, string ownerPassword) =>
        Create(OfficeWorkflowOperation.RemovePdfProtection, inputPath).WithPdfPassword(ownerPassword).WithPdfOwnerPassword(ownerPassword);

    /// <summary>Creates a signed PDF copy with captured appearance settings and a caller-owned signer and verifier.</summary>
    public static OfficeWorkflowBuilder SignPdf(string inputPath, OfficeIMO.Pdf.IPdfExternalSigner signer,
        OfficeIMO.Pdf.PdfExternalSignatureOptions options, OfficeIMO.Pdf.IPdfSignatureCryptographyProvider verifier) =>
        new(new OfficeWorkflowRequest { Operation = OfficeWorkflowOperation.SignPdf, InputPath = inputPath,
            OutputSigner = signer ?? throw new ArgumentNullException(nameof(signer)),
            OutputSignatureOptions = OfficeWorkflowRunner.SnapshotSignatureOptions(options) ?? throw new ArgumentNullException(nameof(options)),
            OutputSignatureValidator = verifier ?? throw new ArgumentNullException(nameof(verifier)) });

    /// <summary>Runs an explicitly constructed request through the default local runner.</summary>
    public static Task<OfficeWorkflowResult> RunAsync(
        OfficeWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        new OfficeWorkflowRunner().RunAsync(request, progress, cancellationToken);

    private static OfficeWorkflowBuilder Create(OfficeWorkflowOperation operation, string inputPath) {
        if ((inputPath == null || string.IsNullOrWhiteSpace(inputPath))) throw new ArgumentException("Input path cannot be empty.", nameof(inputPath));
        return new OfficeWorkflowBuilder(new OfficeWorkflowRequest {
            Operation = operation,
            InputPath = inputPath
        });
    }
}

/// <summary>Fluent configuration for one local document workflow.</summary>
public sealed class OfficeWorkflowBuilder {
    private readonly OfficeWorkflowRequest _request;

    internal OfficeWorkflowBuilder(OfficeWorkflowRequest request) {
        _request = request ?? throw new ArgumentNullException(nameof(request));
    }



    /// <summary>Sets the output artifact path.</summary>
    public OfficeWorkflowBuilder To(string outputPath) {
        if ((outputPath == null || string.IsNullOrWhiteSpace(outputPath))) throw new ArgumentException("Output path cannot be empty.", nameof(outputPath));
        _request.OutputPath = outputPath;
        return this;
    }

    /// <summary>Sets the comparison input for a compare workflow.</summary>
    public OfficeWorkflowBuilder Against(string comparisonPath) {
        if ((comparisonPath == null || string.IsNullOrWhiteSpace(comparisonPath))) throw new ArgumentException("Comparison path cannot be empty.", nameof(comparisonPath));
        _request.ComparisonPath = comparisonPath;
        return this;
    }

    /// <summary>Sets the cross-format output profile.</summary>
    public OfficeWorkflowBuilder WithProfile(OfficeConversionProfile profile) {
        _request.OutputProfile = profile;
        return this;
    }



    /// <summary>Sets destination conflict behavior.</summary>
    public OfficeWorkflowBuilder OnConflict(OfficePublicationPolicy policy) {
        _request.ConflictPolicy = policy;
        return this;
    }

    /// <summary>Sets the PDF password used for the primary input.</summary>
    public OfficeWorkflowBuilder WithPdfPassword(string? password) {
        _request.PdfPassword = password;
        return this;
    }

    /// <summary>Sets current owner authorization for replacing or removing PDF protection.</summary>
    public OfficeWorkflowBuilder WithPdfOwnerPassword(string? password) {
        _request.PdfOwnerPassword = password;
        return this;
    }



    /// <summary>Sets the PDF password used for the comparison input.</summary>
    public OfficeWorkflowBuilder WithComparisonPdfPassword(string? password) {
        _request.ComparisonPdfPassword = password;
        return this;
    }

    /// <summary>Sets bounded input and output sizes for the workflow.</summary>
    public OfficeWorkflowBuilder WithLimits(long maximumInputBytes, long maximumOutputBytes) {
        _request.Limits = new OfficeOperationLimits {
            MaximumInputBytes = maximumInputBytes,
            MaximumOutputBytes = maximumOutputBytes
        };
        return this;
    }

    /// <summary>Sets the caller-visible request identifier.</summary>
    public OfficeWorkflowBuilder WithId(string id) {
        if ((id == null || string.IsNullOrWhiteSpace(id))) throw new ArgumentException("Request id cannot be empty.", nameof(id));
        _request.Id = id;
        return this;
    }

    /// <summary>Selects independently ordered page scopes for ordinal rendered comparison. Null selects all pages on that side.</summary>
    public OfficeWorkflowBuilder ComparePages(OfficeIMO.Pdf.PdfPageSelector? expectedPages, OfficeIMO.Pdf.PdfPageSelector? actualPages) {
        if (_request.Operation != OfficeWorkflowOperation.Compare) throw new InvalidOperationException("Page comparison scopes require Compare.");
        _request.ComparisonExpectedPages = expectedPages;
        _request.ComparisonActualPages = actualPages;
        return this;
    }

    /// <summary>Builds an independent request snapshot, inferring a locally executable conversion route from file extensions when possible.</summary>
    public OfficeWorkflowRequest Build() {
        OfficeOperationLimits limits = _request.Limits ?? throw new InvalidOperationException("Workflow limits cannot be null.");
        return new OfficeWorkflowRequest {
            Id = _request.Id,
            Operation = _request.Operation,
            PageNumbers = _request.PageNumbers?.ToArray(),
            OutputEncryption = _request.OutputEncryption?.Clone(),
            PdfOwnerPassword = _request.PdfOwnerPassword,
            OutputSigner = _request.OutputSigner,
            OutputSignatureOptions = OfficeWorkflowRunner.SnapshotSignatureOptions(_request.OutputSignatureOptions),
            OutputSignatureValidator = _request.OutputSignatureValidator,
            InputPath = _request.InputPath,
            ComparisonPath = _request.ComparisonPath,
            ComparisonExpectedPages = _request.ComparisonExpectedPages,
            ComparisonActualPages = _request.ComparisonActualPages,
            ConversionRouteId = _request.ConversionRouteId,
            ConversionOptions = _request.ConversionOptions?.Clone(),
            RegisteredConversionSettings = _request.RegisteredConversionSettings?.Snapshot(),
            OutputPath = _request.OutputPath,
            ConflictPolicy = _request.ConflictPolicy,
            OutputProfile = _request.OutputProfile,
            PdfPassword = _request.PdfPassword,
            ComparisonPdfPassword = _request.ComparisonPdfPassword,
            Limits = new OfficeOperationLimits {
                MaximumInputBytes = limits.MaximumInputBytes,
                MaximumOutputBytes = limits.MaximumOutputBytes,
                MaximumXmlCharactersInPart = limits.MaximumXmlCharactersInPart
            }
        };
    }

    /// <summary>Runs this workflow through the supplied runner or the default local runner.</summary>
    public Task<OfficeWorkflowResult> RunAsync(
        IOfficeWorkflowRunner? runner = null,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        (runner ?? new OfficeWorkflowRunner()).RunAsync(Build(), progress, cancellationToken);
}
