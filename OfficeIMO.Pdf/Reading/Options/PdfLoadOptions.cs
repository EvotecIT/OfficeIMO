using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>
/// Options for controlling PDF reading/decoding behavior.
/// </summary>
public sealed partial class PdfLoadOptions {
    /// <summary>Creates default read settings with an independent limits graph.</summary>
    internal static PdfLoadOptions Default => new PdfLoadOptions();

    /// <summary>Creates strict parsing settings with bounded budgets for documents from outside.</summary>
    public static PdfLoadOptions UntrustedDefaults => new PdfLoadOptions {
        ParsingMode = PdfParsingMode.Strict,
        Limits = new PdfReadLimits {
            MaxInputBytes = 64L * 1024L * 1024L,
            MaxIndirectObjects = 100_000,
            MaxRawStreamBytes = 32 * 1024 * 1024,
            MaxDecodedStreamBytes = 64 * 1024 * 1024,
            MaxTotalDecodedStreamBytes = 128L * 1024L * 1024L
        }
    };

    /// <summary>Structural parsing policy. Lenient recovery is the compatibility default and always produces a repair report.</summary>
    public PdfParsingMode ParsingMode { get; init; } = PdfParsingMode.Lenient;

    /// <summary>Resource budgets for object scanning and raw stream allocation.</summary>
    public PdfReadLimits Limits { get; init; } = new PdfReadLimits();

    /// <summary>
    /// Password used to open encrypted PDFs. The same value is tried as user and owner password for Standard security handler files.
    /// For Standard security revisions 5 and 6, passwords longer than 4096 UTF-16 characters are rejected before normalization.
    /// </summary>
    public string? Password { get; init; }

    /// <summary>
    /// Optional synchronous AES provider used to open Standard-security PDFs when platform AES-CBC is unavailable.
    /// </summary>
    public IOfficeAesCryptographyProvider? AesCryptographyProvider { get; init; }

    /// <summary>
    /// Controls whether authenticated user-password permission restrictions are enforced.
    /// Ignoring restrictions still requires the PDF to be successfully decrypted with a valid password.
    /// </summary>
    public PdfPermissionPolicy PermissionPolicy { get; init; } = PdfPermissionPolicy.Enforce;
    /// <summary>Prefer decoding via ToUnicode CMap when available. Default: true.</summary>
    public bool PreferToUnicode { get; init; } = true;
    /// <summary>Fallback to WinAnsi (Windows-1252) when no ToUnicode is present. Default: true.</summary>
    public bool UseWinAnsiFallback { get; init; } = true;
    /// <summary>Adjust X position using TJ kerning values (thousandths of font size). Default: true.</summary>
    public bool AdjustKerningFromTJ { get; init; } = true;
    /// <summary>
    /// Includes text inside PDF artifact marked-content sequences, such as page headers, footers,
    /// and chart decorations. Default: false, which returns logical document text.
    /// </summary>
    public bool IncludeArtifactText { get; init; }

    internal static PdfLoadOptions Resolve(PdfLoadOptions? options) {
        PdfLoadOptions effective = options ?? Default;
        Guard.NotNull(effective.Limits, nameof(Limits));
        effective.Limits.Validate();
        if (effective.Profile < PdfReadProfile.Fast || effective.Profile > PdfReadProfile.Structured)
            throw new ArgumentOutOfRangeException(nameof(options), effective.Profile, "Unknown PDF read profile.");
        Guard.NotNull(effective.LayoutOptions, nameof(LayoutOptions));
        return effective;
    }

    internal static PdfLoadOptions WithMinimumInputBytes(PdfLoadOptions? options, long minimumInputBytes) {
        PdfLoadOptions effective = Resolve(options);
        if (minimumInputBytes <= effective.Limits.MaxInputBytes) return effective;
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Profile = effective.Profile,
            IncludeImages = effective.IncludeImages,
            IncludeFormWidgets = effective.IncludeFormWidgets,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline),
            Limits = effective.Limits.WithMinimumInputBytes(minimumInputBytes),
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfLoadOptions WithGeneratedOutputGrowth(
        PdfLoadOptions? options,
        int minimumIndirectObjects,
        PdfGeneratedOutputGrowth growth) {
        PdfLoadOptions effective = Resolve(options);
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Profile = effective.Profile,
            IncludeImages = effective.IncludeImages,
            IncludeFormWidgets = effective.IncludeFormWidgets,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline),
            Limits = effective.Limits.WithGeneratedOutput(effective.Limits.MaxInputBytes, minimumIndirectObjects, growth),
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfLoadOptions ForGeneratedOutput(
        PdfLoadOptions? sourceOptions,
        byte[] sourcePdf,
        byte[] outputPdf,
        PdfGeneratedOutputGrowth growth = default) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(outputPdf, nameof(outputPdf));
        PdfLoadOptions source = Resolve(sourceOptions);
        PdfStructuralMarkerCounts sourceMarkers = PdfSyntax.InspectStructuralMarkers(sourcePdf, source.Limits);
        PdfStructuralMarkerCounts outputMarkers = PdfSyntax.InspectStructuralMarkers(outputPdf, source.Limits);
        int sourceHeaders = sourceMarkers.IndirectObjectHeaders;
        int outputHeaders = outputMarkers.IndirectObjectHeaders;
        int addedHeaders = Math.Max(0, outputHeaders - sourceHeaders);
        int reservedObjects = source.Limits.MaxIndirectObjects > int.MaxValue - addedHeaders
            ? int.MaxValue
            : source.Limits.MaxIndirectObjects + addedHeaders;
        int addedStartXrefMarkers = Math.Max(0, outputMarkers.StartXrefMarkers - sourceMarkers.StartXrefMarkers);
        PdfGeneratedOutputGrowth effectiveGrowth = new PdfGeneratedOutputGrowth(
            additionalRevisions: Math.Max(growth.AdditionalRevisions, addedStartXrefMarkers),
            additionalFormFields: growth.AdditionalFormFields,
            additionalAnnotationsPerPage: growth.AdditionalAnnotationsPerPage,
            minimumRawStreamBytes: Math.Max(growth.MinimumRawStreamBytes, outputMarkers.MaximumRawStreamBytes),
            minimumDecodedStreamBytes: growth.MinimumDecodedStreamBytes,
            additionalTotalDecodedStreamBytes: growth.AdditionalTotalDecodedStreamBytes,
            additionalPageContentBytes: growth.AdditionalPageContentBytes,
            additionalRetainedContentBytes: growth.AdditionalRetainedContentBytes,
            additionalDecodedTextCharacters: growth.AdditionalDecodedTextCharacters,
            minimumObjectCharacters: Math.Max(growth.MinimumObjectCharacters, outputMarkers.MaximumObjectCharacters),
            minimumTokensPerObject: Math.Max(growth.MinimumTokensPerObject, outputMarkers.MaximumObjectCharacters),
            minimumObjectNestingDepth: growth.MinimumObjectNestingDepth,
            additionalContentOperations: growth.AdditionalContentOperations,
            additionalContentOperands: growth.AdditionalContentOperands,
            additionalContentNestingDepth: growth.AdditionalContentNestingDepth,
            additionalPrintProductionContexts: Math.Max(growth.AdditionalPrintProductionContexts, addedHeaders),
            additionalPrintProductionOperations: growth.AdditionalPrintProductionOperations);

        return new PdfLoadOptions {
            ParsingMode = source.ParsingMode,
            Profile = source.Profile,
            IncludeImages = source.IncludeImages,
            IncludeFormWidgets = source.IncludeFormWidgets,
            PageSelection = source.PageSelection,
            LayoutOptions = CloneLayoutOptions(source.LayoutOptions),
            Pipeline = ClonePipelineOptions(source.Pipeline),
            Limits = source.Limits.WithGeneratedOutput(outputPdf.LongLength, Math.Max(outputHeaders, reservedObjects), effectiveGrowth),
            Password = source.Password,
            AesCryptographyProvider = source.AesCryptographyProvider,
            PermissionPolicy = source.PermissionPolicy,
            PreferToUnicode = source.PreferToUnicode,
            UseWinAnsiFallback = source.UseWinAnsiFallback,
            AdjustKerningFromTJ = source.AdjustKerningFromTJ,
            IncludeArtifactText = source.IncludeArtifactText
        };
    }

    internal static PdfLoadOptions ForComposedOutput(
        PdfLoadOptions? primaryOptions,
        IEnumerable<PdfLoadOptions> sourceOptions,
        long minimumInputBytes,
        int minimumIndirectObjects) {
        Guard.NotNull(sourceOptions, nameof(sourceOptions));
        PdfLoadOptions primary = Resolve(primaryOptions);
        PdfLoadOptions[] sources = sourceOptions.Select(Resolve).ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one source read-options instance is required for composed output.", nameof(sourceOptions));
        }

        return new PdfLoadOptions {
            ParsingMode = primary.ParsingMode,
            Profile = primary.Profile,
            IncludeImages = primary.IncludeImages,
            IncludeFormWidgets = primary.IncludeFormWidgets,
            PageSelection = primary.PageSelection,
            LayoutOptions = CloneLayoutOptions(primary.LayoutOptions),
            Pipeline = ClonePipelineOptions(primary.Pipeline),
            Limits = PdfReadLimits.ForComposedOutput(
                sources.Select(static source => source.Limits).ToArray(),
                minimumInputBytes,
                minimumIndirectObjects),
            PermissionPolicy = primary.PermissionPolicy,
            PreferToUnicode = primary.PreferToUnicode,
            UseWinAnsiFallback = primary.UseWinAnsiFallback,
            AdjustKerningFromTJ = primary.AdjustKerningFromTJ,
            IncludeArtifactText = primary.IncludeArtifactText
        };
    }

    internal static PdfLoadOptions WithMaximumContainerEntries(
        PdfLoadOptions? options,
        int maximumContainerEntries,
        long? maximumDecodedStreamBytes = null,
        long? maximumTotalDecodedStreamBytes = null,
        long? maximumTotalAttachmentBytes = null,
        bool preserveExistingDecodedStreamLimit = true,
        long? maximumRawStreamBytes = null,
        bool preserveExistingRawStreamLimit = true) {
        PdfLoadOptions effective = Resolve(options);
        if (maximumContainerEntries <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumContainerEntries), maximumContainerEntries, "Maximum container entries must be positive.");
        }
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Profile = effective.Profile,
            IncludeImages = effective.IncludeImages,
            IncludeFormWidgets = effective.IncludeFormWidgets,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline),
            Limits = effective.Limits.WithMaximumContainerEntries(
                maximumContainerEntries,
                maximumDecodedStreamBytes,
                maximumTotalDecodedStreamBytes,
                maximumTotalAttachmentBytes,
                preserveExistingDecodedStreamLimit,
                maximumRawStreamBytes,
                preserveExistingRawStreamLimit),
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfLoadOptions WithPassword(PdfLoadOptions? options, string? password) {
        PdfLoadOptions effective = Resolve(options);
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Profile = effective.Profile,
            IncludeImages = effective.IncludeImages,
            IncludeFormWidgets = effective.IncludeFormWidgets,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline),
            Limits = effective.Limits,
            Password = password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfLoadOptions WithAesCryptographyProvider(
        PdfLoadOptions? options,
        IOfficeAesCryptographyProvider? aesCryptographyProvider) {
        PdfLoadOptions effective = Resolve(options);
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Profile = effective.Profile,
            IncludeImages = effective.IncludeImages,
            IncludeFormWidgets = effective.IncludeFormWidgets,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline),
            Limits = effective.Limits,
            Password = effective.Password,
            AesCryptographyProvider = aesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfLoadOptions WithArtifactText(PdfLoadOptions? options) {
        PdfLoadOptions effective = Resolve(options);
        if (effective.IncludeArtifactText) return effective;
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Profile = effective.Profile,
            IncludeImages = effective.IncludeImages,
            IncludeFormWidgets = effective.IncludeFormWidgets,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline),
            Limits = effective.Limits,
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = true
        };
    }
}
