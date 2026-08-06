using OfficeIMO.Core.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    private static readonly string[] SupportedPowerPointConversionExtensions = PowerPointFormatCatalog.All
        .Select(format => format.Extension)
        .ToArray();

    /// <summary>Converts a PowerPoint file and returns format and feature-level fidelity diagnostics.</summary>
    public static PowerPointPresentationConversionResult Convert(
        string sourcePath,
        string destinationPath,
        PowerPointPresentationConversionOptions? options = null) =>
        ConvertAsync(sourcePath, destinationPath, options).GetAwaiter().GetResult();

    /// <summary>Asynchronously converts a PowerPoint file and returns feature-level fidelity diagnostics.</summary>
    public static async Task<PowerPointPresentationConversionResult> ConvertAsync(
        string sourcePath,
        string destinationPath,
        PowerPointPresentationConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new PowerPointPresentationConversionOptions();
        OfficeFileConversion.Paths paths = OfficeFileConversion.ValidatePaths(
            sourcePath,
            destinationPath,
            SupportedPowerPointConversionExtensions,
            "PowerPoint presentation");

        PowerPointLoadOptions loadOptions = CreateConversionLoadOptions(options.LoadOptions);
        using PowerPointPresentation presentation = await LoadAsync(paths.Source, loadOptions, cancellationToken)
            .ConfigureAwait(false);
        presentation.SignatureMutationPolicy = options.SignatureMutationPolicy;
        OfficeFormatDescriptor sourceDescriptor = presentation.SourceFormatDescriptor;
        OfficeFormatDescriptor destinationDescriptor = PowerPointFormatCatalog.GetByExtension(paths.Destination);
        PowerPointFileFormat sourceFormat = presentation.SourceFormat;
        PowerPointFileFormat destinationFormat = PowerPointPresentationLoadRouting.GetFormat(paths.Destination);
        OfficeCompatibilityMode compatibilityMode = GetCompatibilityMode(options);
        bool allowsLoss = AllowsLoss(options, compatibilityMode);
        List<OfficeConversionDiagnostic> diagnostics = CreateOfficeConversionDiagnostics(
            presentation,
            paths.Source,
            sourceDescriptor,
            destinationDescriptor,
            options,
            compatibilityMode,
            allowsLoss,
            out bool embedSourceCarrier);
        PowerPointPresentationConversionResult assessment = CreatePowerPointConversionResult(
            paths,
            sourceFormat,
            destinationFormat,
            sourceDescriptor,
            destinationDescriptor,
            compatibilityMode,
            diagnostics,
            outputCreated: false,
            replacedExistingFile: false);

        if (sourceDescriptor.Equals(destinationDescriptor)) {
            throw new PowerPointPresentationConversionException(
                OfficeConversionFailureReason.SameFormat,
                assessment,
                $"The source is already {sourceDescriptor.Id}. Convert requires a different concrete source and destination format.");
        }

        if (diagnostics.Any(diagnostic => diagnostic.RepresentsDataLoss
                && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked)) {
            throw new PowerPointPresentationConversionException(
                OfficeConversionFailureReason.DataLossBlocked,
                assessment,
                $"PowerPoint conversion is blocked because {diagnostics.Count(diagnostic => diagnostic.RepresentsDataLoss && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked)} feature(s) have no representation accepted by the selected compatibility policy. Inspect Result.Report.Compatibility or select an explicit fallback policy.");
        }

        if (assessment.Report.Compatibility.HasBlockedFeatures) {
            throw new PowerPointPresentationConversionException(
                OfficeConversionFailureReason.DestinationFeatureUnsupported,
                assessment,
                $"The requested destination {destinationDescriptor.Id} is classified but is not a supported native write target. Inspect Result.Report.Compatibility for details.");
        }

        if (File.Exists(paths.Destination)
            && options.FileConflictPolicy == OfficeConversionFileConflictPolicy.FailIfExists) {
            throw new PowerPointPresentationConversionException(
                OfficeConversionFailureReason.DestinationExists,
                assessment,
                $"The destination file '{paths.Destination}' already exists. Set FileConflictPolicy to Replace to replace it atomically.");
        }

        EnsureDestinationFileWritable(paths.Destination);
        OfficeFileConversion.EnsureDestinationDirectory(paths.Destination);
        string stagingPath = OfficeFileCommit.CreateStagingPath(paths.Destination);
        try {
            try {
                byte[]? sourceBytes = embedSourceCarrier
                    ? await OfficeFileConversion.ReadAllBytesAsync(paths.Source, cancellationToken).ConfigureAwait(false)
                    : null;
                await presentation.SaveAsync(
                    stagingPath,
                    CreateConversionSaveOptions(options.SaveOptions, allowsLoss),
                    cancellationToken).ConfigureAwait(false);
                if (embedSourceCarrier) {
                    byte[] destinationBytes = await OfficeFileConversion.ReadAllBytesAsync(
                        stagingPath,
                        cancellationToken).ConfigureAwait(false);
                    byte[] carriedBytes = AttachPowerPointSourceCarrier(
                        destinationBytes,
                        destinationDescriptor,
                        sourceDescriptor.Id,
                        Path.GetFileName(paths.Source),
                        compatibilityMode,
                        sourceBytes ?? throw new InvalidOperationException("Embedded-source conversion requires source bytes."));
                    await OfficeFileCommit.WriteAllBytesAsync(
                        stagingPath,
                        carriedBytes,
                        cancellationToken: cancellationToken).ConfigureAwait(false);
                }
            } catch (NotSupportedException exception) {
                diagnostics.Add(new OfficeConversionDiagnostic(
                    "PowerPoint.DestinationFeatureUnsupported",
                    OfficeConversionDiagnosticCategory.DestinationFormat,
                    OfficeConversionDiagnosticSeverity.Error,
                    exception.Message,
                    OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.Semantic,
                    representsDataLoss: false));
                throw new PowerPointPresentationConversionException(
                    OfficeConversionFailureReason.DestinationFeatureUnsupported,
                    CreatePowerPointConversionResult(
                        paths,
                        sourceFormat,
                        destinationFormat,
                        sourceDescriptor,
                        destinationDescriptor,
                        compatibilityMode,
                        diagnostics,
                        false,
                        false),
                    $"The presentation contains content that cannot be written as {destinationDescriptor.Id}. See Result.Report.Compatibility for the feature-level decision.",
                    exception);
            }

            bool replacesExistingFile = File.Exists(paths.Destination);
            try {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeFileCommit.CommitTemporaryFile(
                    stagingPath,
                    paths.Destination,
                    options.FileConflictPolicy == OfficeConversionFileConflictPolicy.Replace
                        ? OfficeFileCommit.ConflictPolicy.Replace
                        : OfficeFileCommit.ConflictPolicy.FailIfExists);
                stagingPath = string.Empty;
            } catch (IOException exception) when (
                options.FileConflictPolicy == OfficeConversionFileConflictPolicy.FailIfExists
                && File.Exists(paths.Destination)) {
                throw new PowerPointPresentationConversionException(
                    OfficeConversionFailureReason.DestinationExists,
                    assessment,
                    $"The destination file '{paths.Destination}' was created while conversion was running and was not replaced.",
                    exception);
            }

            return CreatePowerPointConversionResult(
                paths,
                sourceFormat,
                destinationFormat,
                sourceDescriptor,
                destinationDescriptor,
                compatibilityMode,
                diagnostics,
                true,
                replacesExistingFile);
        } finally {
            OfficeFileCommit.DeleteIfExists(stagingPath);
        }
    }

    private static List<OfficeConversionDiagnostic> CreateOfficeConversionDiagnostics(
        PowerPointPresentation presentation,
        string sourcePath,
        OfficeFormatDescriptor sourceDescriptor,
        OfficeFormatDescriptor destinationDescriptor,
        PowerPointPresentationConversionOptions options,
        OfficeCompatibilityMode compatibilityMode,
        bool allowsLoss,
        out bool embedSourceCarrier) {
        var diagnostics = new List<OfficeConversionDiagnostic>();
        bool preserveLossySource = allowsLoss
            && (options.EmbedSourceWhenLossy
                || compatibilityMode == OfficeCompatibilityMode.PreservationOnly);
        bool permitsVisualFallback = compatibilityMode is OfficeCompatibilityMode.PreferVisual
            or OfficeCompatibilityMode.BestEffort
            or OfficeCompatibilityMode.PreservationOnly;
        OfficeFormatDescriptor declaredSource = PowerPointFormatCatalog.GetByExtension(sourcePath);
        if (!declaredSource.Equals(sourceDescriptor)) {
            diagnostics.Add(new OfficeConversionDiagnostic(
                "PowerPoint.SourceExtensionMismatch",
                OfficeConversionDiagnosticCategory.SourceFormat,
                OfficeConversionDiagnosticSeverity.Warning,
                $"The source extension declares {declaredSource.Id}, but its package declares {sourceDescriptor.Id}. Package content was used.",
                OfficeCompatibilityState.Equivalent,
                OfficeCompatibilityImpact.None,
                representsDataLoss: false));
        }

        if (sourceDescriptor.Generation == OfficeFormatGeneration.Legacy
            && destinationDescriptor.Generation == OfficeFormatGeneration.Modern) {
            foreach (LegacyPptImportDiagnostic diagnostic in presentation.LegacyPptImportDiagnostics) {
                bool encryptionRemoved = string.Equals(
                    diagnostic.Code,
                    "PPT-ENCRYPTION-DECRYPTED",
                    StringComparison.Ordinal);
                bool representsLoss = encryptionRemoved
                    || diagnostic.Severity == LegacyPptDiagnosticSeverity.Warning
                    || diagnostic.Severity == LegacyPptDiagnosticSeverity.Error;
                OfficeCompatibilityState state = representsLoss
                    ? GetPowerPointGenericLossState(compatibilityMode, preserveLossySource)
                    : OfficeCompatibilityState.Equivalent;
                diagnostics.Add(new OfficeConversionDiagnostic(
                    diagnostic.Code,
                    representsLoss
                        ? OfficeConversionDiagnosticCategory.DataLoss
                        : OfficeConversionDiagnosticCategory.SourceFormat,
                    state == OfficeCompatibilityState.Blocked
                        ? OfficeConversionDiagnosticSeverity.Error
                        : diagnostic.Severity == LegacyPptDiagnosticSeverity.Error
                        ? OfficeConversionDiagnosticSeverity.Error
                        : diagnostic.Severity == LegacyPptDiagnosticSeverity.Warning
                            ? OfficeConversionDiagnosticSeverity.Warning
                            : OfficeConversionDiagnosticSeverity.Information,
                    encryptionRemoved
                        ? "The legacy presentation was decrypted for import, but the requested modern destination is not password-encrypted. Save the converted artifact with an explicit encryption API if confidentiality must continue."
                        : diagnostic.Message,
                    state,
                    representsLoss
                        ? encryptionRemoved
                            ? OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier
                            : OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability
                        : OfficeCompatibilityImpact.None,
                    representsLoss,
                    diagnostic.StreamOffset.HasValue ? $"PowerPoint Document@0x{diagnostic.StreamOffset.Value:X}" : null));
            }
        }

        if (destinationDescriptor.Generation == OfficeFormatGeneration.Legacy) {
            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            foreach (LegacyPptWriteFinding finding in preflight.Findings) {
                OfficeCompatibilityState state = GetLegacyPptFindingState(
                    finding,
                    compatibilityMode,
                    permitsVisualFallback,
                    preserveLossySource);
                diagnostics.Add(new OfficeConversionDiagnostic(
                    finding.Code,
                    OfficeConversionDiagnosticCategory.DataLoss,
                    state == OfficeCompatibilityState.Blocked
                        ? OfficeConversionDiagnosticSeverity.Error
                        : OfficeConversionDiagnosticSeverity.Warning,
                    finding.Description,
                    state,
                    GetLegacyPptFindingImpact(finding, state),
                    representsDataLoss: true,
                    CreatePowerPointSourceLocation(finding)));
            }
        }

        bool hasVba = presentation._document?.PresentationPart?.VbaProjectPart != null;
        if (hasVba && !destinationDescriptor.IsMacroEnabled) {
            OfficeCompatibilityState macroState = GetPowerPointGenericLossState(
                compatibilityMode,
                preserveLossySource);
            diagnostics.Add(new OfficeConversionDiagnostic(
                "PowerPoint.VbaProject.Removed",
                OfficeConversionDiagnosticCategory.DataLoss,
                macroState == OfficeCompatibilityState.Blocked
                    ? OfficeConversionDiagnosticSeverity.Error
                    : OfficeConversionDiagnosticSeverity.Warning,
                $"The source contains VBA, but {destinationDescriptor.Extension} cannot carry a VBA project.",
                macroState,
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Security,
                representsDataLoss: true,
                sourceLocation: "ppt/vbaProject.bin"));
        }

        PowerPointSignatureReport signatures = presentation.InspectSignatures();
        bool legacySignatureCanRemainByteExact = sourceDescriptor.Generation == OfficeFormatGeneration.Legacy
            && destinationDescriptor.Generation == OfficeFormatGeneration.Legacy;
        if (signatures.HasSignatureMetadata && !legacySignatureCanRemainByteExact) {
            bool signatureRewriteAllowed = options.SignatureMutationPolicy
                != OfficeSignatureMutationPolicy.BlockSave;
            OfficeCompatibilityState signatureState = signatureRewriteAllowed
                ? GetPowerPointGenericLossState(compatibilityMode, preserveLossySource)
                : OfficeCompatibilityState.Blocked;
            diagnostics.Add(new OfficeConversionDiagnostic(
                "PowerPoint.DigitalSignature.Invalidated",
                OfficeConversionDiagnosticCategory.DataLoss,
                signatureState == OfficeCompatibilityState.Blocked
                    ? OfficeConversionDiagnosticSeverity.Error
                    : OfficeConversionDiagnosticSeverity.Warning,
                signatureRewriteAllowed
                    ? "Rewriting the presentation invalidates its existing digital signature. Preserved signature markup must no longer be trusted."
                    : "The source carries digital-signature metadata and the configured mutation policy blocks conversion because rewriting the presentation can invalidate the signature.",
                signatureState,
                OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier,
                representsDataLoss: true,
                sourceLocation: "_xmlsignatures"));
        }


        if (destinationDescriptor.Generation == OfficeFormatGeneration.Legacy
            && string.Equals(destinationDescriptor.Extension, ".ppa", StringComparison.Ordinal)) {
            diagnostics.Add(new OfficeConversionDiagnostic(
                "PowerPoint.LegacyDestination.NotWritable",
                OfficeConversionDiagnosticCategory.DestinationFormat,
                OfficeConversionDiagnosticSeverity.Error,
                ".ppa is classified for import and reporting, but native binary add-in authoring is not yet a supported write target.",
                OfficeCompatibilityState.Blocked,
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Carrier,
                representsDataLoss: false));
        }

        bool hasBlockedFeature = diagnostics.Any(diagnostic =>
            diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked);
        bool hasLossyFeature = diagnostics.Any(diagnostic => diagnostic.RepresentsDataLoss);
        embedSourceCarrier = preserveLossySource
            && !hasBlockedFeature
            && (compatibilityMode == OfficeCompatibilityMode.PreservationOnly || hasLossyFeature);
        if (!hasBlockedFeature && (hasLossyFeature || compatibilityMode == OfficeCompatibilityMode.PreservationOnly)) {
            AddPowerPointSourceCarrierDiagnostic(
                diagnostics,
                embedSourceCarrier,
                hasVba);
        }

        return diagnostics
            .GroupBy(diagnostic => string.Join("\0", diagnostic.Code, diagnostic.Message, diagnostic.SourceLocation), StringComparer.Ordinal)
            .Select(group => group.First())
            .ToList();
    }

    private static OfficeCompatibilityState GetLegacyPptFindingState(
        LegacyPptWriteFinding finding,
        OfficeCompatibilityMode mode,
        bool permitsVisualFallback,
        bool preserveLossySource) {
        if (finding.Feature == LegacyPptFeature.Charts
            || finding.Feature == LegacyPptFeature.SmartArt
            || finding.Feature == LegacyPptFeature.Tables) {
            return permitsVisualFallback
                ? OfficeCompatibilityState.Rasterized
                : OfficeCompatibilityState.Blocked;
        }

        LegacyPptCapability? capability = LegacyPptCapabilityCatalog.Capabilities
            .FirstOrDefault(candidate => candidate.Feature == finding.Feature);
        if (mode != OfficeCompatibilityMode.StrictNative
            && capability?.PptxToBinary == LegacyPptCapabilityState.Converted) {
            return OfficeCompatibilityState.Approximated;
        }

        return GetPowerPointGenericLossState(mode, preserveLossySource);
    }

    private static OfficeCompatibilityState GetPowerPointGenericLossState(
        OfficeCompatibilityMode mode,
        bool embedSource) => mode switch {
            OfficeCompatibilityMode.BestEffort => embedSource
                ? OfficeCompatibilityState.EmbeddedSource
                : OfficeCompatibilityState.Dropped,
            OfficeCompatibilityMode.PreservationOnly => OfficeCompatibilityState.EmbeddedSource,
            _ => OfficeCompatibilityState.Blocked
        };

    private static OfficeCompatibilityImpact GetLegacyPptFindingImpact(
        LegacyPptWriteFinding finding,
        OfficeCompatibilityState state) {
        OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Editability;
        if (state == OfficeCompatibilityState.Rasterized) impact |= OfficeCompatibilityImpact.Behavioral;
        if (state == OfficeCompatibilityState.Dropped) impact |= OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Carrier;
        if (finding.Feature == LegacyPptFeature.VbaProjects
            || finding.Feature == LegacyPptFeature.ActiveX
            || finding.Feature == LegacyPptFeature.EmbeddedOle
            || finding.Feature == LegacyPptFeature.LinkedOle) {
            impact |= OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security;
        }
        return impact;
    }

    private static string? CreatePowerPointSourceLocation(LegacyPptWriteFinding finding) {
        if (!finding.SlideIndex.HasValue) return null;
        return finding.ShapeIndex.HasValue
            ? $"slide:{finding.SlideIndex.Value + 1}/shape:{finding.ShapeIndex.Value + 1}"
            : $"slide:{finding.SlideIndex.Value + 1}";
    }

    private static PowerPointLoadOptions CreateConversionLoadOptions(PowerPointLoadOptions? source) {
        LegacyPptImportOptions? legacy = source?.LegacyPptImportOptions;
        return new PowerPointLoadOptions {
            AccessMode = DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit,
            PackageSecurity = source?.PackageSecurity,
            OpenSettings = source?.OpenSettings,
            MaxInputBytes = source == null
                ? 512L * 1024L * 1024L
                : source.MaxInputBytes,
            LegacyPptImportOptions = legacy == null
                ? new LegacyPptImportOptions { ReportUnsupportedContent = true }
                : new LegacyPptImportOptions {
                    MaxInputBytes = legacy.MaxInputBytes,
                    ReportUnsupportedContent = true,
                    MaxRecordCount = legacy.MaxRecordCount,
                    MaxRecordDepth = legacy.MaxRecordDepth,
                    MaxMasterCount = legacy.MaxMasterCount,
                    MaxConnectorRuleCount = legacy.MaxConnectorRuleCount,
                    MaxCommentCount = legacy.MaxCommentCount,
                    MaxTextStyle9EntryCount = legacy.MaxTextStyle9EntryCount,
                    MaxDecodedStorageBytes = legacy.MaxDecodedStorageBytes,
                    Password = legacy.Password
                }
        };
    }

    private static PowerPointSaveOptions CreateConversionSaveOptions(
        PowerPointSaveOptions? source,
        bool allowsLoss) => new() {
        LossPolicy = allowsLoss ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block,
        LegacyPptEncryptionKeySizeBits = source?.LegacyPptEncryptionKeySizeBits ?? 128,
        LegacyPptEncryptDocumentProperties = source?.LegacyPptEncryptDocumentProperties ?? true,
        LegacyPptAllowUnencryptedCompoundStreams = source?
            .LegacyPptAllowUnencryptedCompoundStreams ?? false
    };

    private static OfficeCompatibilityMode GetCompatibilityMode(PowerPointPresentationConversionOptions options) {
        if (options.CompatibilityMode != OfficeCompatibilityMode.StrictNative) return options.CompatibilityMode;
        return options.LossPolicy == OfficeConversionLossPolicy.Allow
            ? OfficeCompatibilityMode.BestEffort
            : OfficeCompatibilityMode.StrictNative;
    }

    private static bool AllowsLoss(
        PowerPointPresentationConversionOptions options,
        OfficeCompatibilityMode mode) => options.LossPolicy == OfficeConversionLossPolicy.Allow
        || mode == OfficeCompatibilityMode.PreferEditable
        || mode == OfficeCompatibilityMode.PreferVisual
        || mode == OfficeCompatibilityMode.BestEffort
        || mode == OfficeCompatibilityMode.PreservationOnly;

    private static byte[] AttachPowerPointSourceCarrier(
        byte[] destinationBytes,
        OfficeFormatDescriptor destinationFormat,
        string sourceFormatId,
        string sourceFileName,
        OfficeCompatibilityMode mode,
        byte[] sourceBytes) => destinationFormat.Encoding == OfficeFormatEncoding.CompoundBinary
        ? OfficeCompatibilitySourceCarrier.AttachToCompound(
            destinationBytes,
            sourceFormatId,
            sourceFileName,
            mode,
            sourceBytes)
        : OfficeCompatibilitySourceCarrier.AttachToPackage(
            destinationBytes,
            sourceFormatId,
            sourceFileName,
            mode,
            sourceBytes);

    private static void AddPowerPointSourceCarrierDiagnostic(
        List<OfficeConversionDiagnostic> diagnostics,
        bool embedded,
        bool hasVba) {
        diagnostics.Add(new OfficeConversionDiagnostic(
            embedded ? "PowerPoint.SourceCarrier.Embedded" : "PowerPoint.SourceCarrier.NotEmbedded",
            OfficeConversionDiagnosticCategory.DataLoss,
            OfficeConversionDiagnosticSeverity.Warning,
            embedded
                ? "The complete original source is retained in an inert, hash-verified OfficeIMO compatibility carrier. It is not executable or editable through the converted presentation model."
                : "The complete original source is not retained. Set EmbedSourceWhenLossy when deliberate byte-level recovery is required.",
            embedded ? OfficeCompatibilityState.EmbeddedSource : OfficeCompatibilityState.Dropped,
            OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability
                | (hasVba
                    ? OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Behavioral
                    : OfficeCompatibilityImpact.None),
            representsDataLoss: !embedded,
            fallbackArtifact: embedded ? OfficeCompatibilitySourceCarrier.PayloadPath : null));
    }

    private static PowerPointPresentationConversionResult CreatePowerPointConversionResult(
        OfficeFileConversion.Paths paths,
        PowerPointFileFormat sourceFormat,
        PowerPointFileFormat destinationFormat,
        OfficeFormatDescriptor sourceDescriptor,
        OfficeFormatDescriptor destinationDescriptor,
        OfficeCompatibilityMode compatibilityMode,
        IReadOnlyList<OfficeConversionDiagnostic> diagnostics,
        bool outputCreated,
        bool replacedExistingFile) => new(
            paths.Source,
            paths.Destination,
            sourceFormat,
            destinationFormat,
            sourceDescriptor,
            destinationDescriptor,
            compatibilityMode,
            diagnostics.ToArray(),
            outputCreated,
            replacedExistingFile);
}
