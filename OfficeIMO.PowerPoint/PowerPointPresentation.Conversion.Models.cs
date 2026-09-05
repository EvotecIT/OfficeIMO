using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>Represents the destination artifact and report produced by a PowerPoint conversion.</summary>
public sealed class PowerPointPresentationConversionResult : IOfficeConversionResult<string, PowerPointPresentationConversionReport> {
    internal PowerPointPresentationConversionResult(
        string sourcePath,
        string destinationPath,
        PowerPointFileFormat sourceFormat,
        PowerPointFileFormat destinationFormat,
        OfficeFormatDescriptor sourceDescriptor,
        OfficeFormatDescriptor destinationDescriptor,
        OfficeCompatibilityMode compatibilityMode,
        IReadOnlyList<OfficeConversionDiagnostic> diagnostics,
        bool outputCreated,
        bool replacedExistingFile) {
        Value = outputCreated ? destinationPath : null;
        Report = new PowerPointPresentationConversionReport(
            sourcePath,
            destinationPath,
            sourceFormat,
            destinationFormat,
            sourceDescriptor,
            destinationDescriptor,
            compatibilityMode,
            diagnostics,
            replacedExistingFile);
    }

    /// <summary>Gets the normalized destination path when committed; otherwise, <see langword="null"/>.</summary>
    public string? Value { get; }

    /// <summary>Gets the immutable conversion assessment.</summary>
    public PowerPointPresentationConversionReport Report { get; }

    /// <summary>Gets whether the conversion reports fidelity loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Gets whether the destination artifact was committed.</summary>
    public bool Succeeded => Value != null;

    /// <summary>Returns the committed path or throws when no artifact was produced.</summary>
    public string RequireValue() => Value
        ?? throw new InvalidOperationException("The PowerPoint conversion did not produce a destination artifact.");

    /// <summary>Returns the committed path only when no loss was reported.</summary>
    public string RequireNoLoss() {
        Report.RequireNoLoss();
        return RequireValue();
    }
}

/// <summary>Describes paths, formats, fidelity decisions, and commit behavior for one conversion.</summary>
public sealed class PowerPointPresentationConversionReport : IOfficeConversionReport {
    internal PowerPointPresentationConversionReport(
        string sourcePath,
        string destinationPath,
        PowerPointFileFormat sourceFormat,
        PowerPointFileFormat destinationFormat,
        OfficeFormatDescriptor sourceDescriptor,
        OfficeFormatDescriptor destinationDescriptor,
        OfficeCompatibilityMode compatibilityMode,
        IReadOnlyList<OfficeConversionDiagnostic> diagnostics,
        bool replacedExistingFile) {
        SourcePath = sourcePath;
        DestinationPath = destinationPath;
        SourceFormat = sourceFormat;
        DestinationFormat = destinationFormat;
        SourceFormatDescriptor = sourceDescriptor;
        DestinationFormatDescriptor = destinationDescriptor;
        Diagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
        Compatibility = new OfficeCompatibilityReport(
            sourceDescriptor,
            destinationDescriptor,
            compatibilityMode,
            Diagnostics.Select(ToCompatibilityFinding));
        ReplacedExistingFile = replacedExistingFile;
    }

    /// <summary>Gets the normalized source path.</summary>
    public string SourcePath { get; }

    /// <summary>Gets the normalized destination path.</summary>
    public string DestinationPath { get; }

    /// <summary>Gets the source's broad physical format.</summary>
    public PowerPointFileFormat SourceFormat { get; }

    /// <summary>Gets the destination's broad physical format.</summary>
    public PowerPointFileFormat DestinationFormat { get; }

    /// <summary>Gets the concrete source format and document kind.</summary>
    public OfficeFormatDescriptor SourceFormatDescriptor { get; }

    /// <summary>Gets the concrete destination format and document kind.</summary>
    public OfficeFormatDescriptor DestinationFormatDescriptor { get; }

    /// <summary>Gets all PowerPoint-specific diagnostics.</summary>
    public IReadOnlyList<OfficeConversionDiagnostic> Diagnostics { get; }

    /// <summary>Gets the shared feature-level fidelity assessment.</summary>
    public OfficeCompatibilityReport Compatibility { get; }

    /// <summary>Gets whether the conversion reports fidelity loss.</summary>
    public bool HasLoss => Compatibility.HasLoss;

    /// <summary>Gets whether an existing destination was replaced.</summary>
    public bool ReplacedExistingFile { get; }

    /// <summary>Throws when the conversion reports loss or a blocked feature.</summary>
    public void RequireNoLoss() => Compatibility.RequireNoLoss();

    private static OfficeCompatibilityFinding ToCompatibilityFinding(OfficeConversionDiagnostic diagnostic) => new(
        diagnostic.Code,
        diagnostic.Category.ToString(),
        diagnostic.Message,
        diagnostic.CompatibilityState,
        diagnostic.Severity switch {
            OfficeConversionDiagnosticSeverity.Warning => OfficeCompatibilitySeverity.Warning,
            OfficeConversionDiagnosticSeverity.Error => OfficeCompatibilitySeverity.Error,
            _ => OfficeCompatibilitySeverity.Information
        },
        diagnostic.CompatibilityImpact,
        diagnostic.RepresentsDataLoss,
        diagnostic.SourceLocation,
        diagnostic.FallbackArtifact);
}

/// <summary>Raised when a validated PowerPoint conversion cannot be completed safely.</summary>
public sealed class PowerPointPresentationConversionException : InvalidOperationException {
    internal PowerPointPresentationConversionException(
        OfficeConversionFailureReason reason,
        PowerPointPresentationConversionResult result,
        string message,
        Exception? innerException = null)
        : base(message, innerException) {
        Reason = reason;
        Result = result;
    }

    /// <summary>Gets the structured failure reason.</summary>
    public OfficeConversionFailureReason Reason { get; }

    /// <summary>Gets the assessment available when conversion was rejected.</summary>
    public PowerPointPresentationConversionResult Result { get; }
}

/// <summary>Controls file-to-file PowerPoint conversion.</summary>
public sealed class PowerPointPresentationConversionOptions {
    /// <summary>Gets or sets how an existing destination is handled.</summary>
    public OfficeConversionFileConflictPolicy FileConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.FailIfExists;

    /// <summary>Gets or sets whether known conversion loss is blocked.</summary>
    public OfficeConversionLossPolicy LossPolicy { get; set; } = OfficeConversionLossPolicy.Block;

    /// <summary>Gets or sets the requested editability, visual-fidelity, and preservation strategy.</summary>
    public OfficeCompatibilityMode CompatibilityMode { get; set; } = OfficeCompatibilityMode.StrictNative;

    /// <summary>
    /// Gets or sets whether the complete original file is retained in an inert, hash-verified
    /// compatibility carrier when an allowed conversion is lossy.
    /// </summary>
    /// <remarks>
    /// Original presentations can contain macros, embedded objects, linked content, or hidden data.
    /// The carrier is not executed by OfficeIMO, but callers should apply the same trust policy they
    /// use for the original file before extracting or opening it. <see cref="OfficeCompatibilityMode.PreservationOnly"/>
    /// enables this setting automatically.
    /// </remarks>
    public bool EmbedSourceWhenLossy { get; set; }

    /// <summary>
    /// Gets or sets how conversion handles existing digital-signature metadata. The safe default blocks
    /// package rewriting; removing or preserving invalidated markup still produces a reported security loss.
    /// </summary>
    public OfficeSignatureMutationPolicy SignatureMutationPolicy { get; set; } =
        OfficeSignatureMutationPolicy.BlockSave;

    /// <summary>Gets or sets optional source load settings.</summary>
    public PowerPointLoadOptions? LoadOptions { get; set; }

    /// <summary>Gets or sets optional destination save settings.</summary>
    public PowerPointSaveOptions? SaveOptions { get; set; }
}
