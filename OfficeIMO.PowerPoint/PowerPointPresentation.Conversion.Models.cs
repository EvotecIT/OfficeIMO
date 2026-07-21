using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>Controls how conversion handles an existing PowerPoint destination.</summary>
public enum PowerPointConversionFileConflictPolicy {
    /// <summary>Reject conversion if the destination exists.</summary>
    FailIfExists,
    /// <summary>Replace an existing destination through an atomic commit.</summary>
    Replace
}

/// <summary>Identifies the purpose of a PowerPoint conversion diagnostic.</summary>
public enum PowerPointConversionDiagnosticCategory {
    /// <summary>Source format detection or extension findings.</summary>
    SourceFormat,
    /// <summary>Content that cannot survive conversion unchanged.</summary>
    DataLoss,
    /// <summary>Destination format or writer findings.</summary>
    DestinationFormat
}

/// <summary>Identifies the severity of a PowerPoint conversion diagnostic.</summary>
public enum PowerPointConversionDiagnosticSeverity {
    /// <summary>Informational finding.</summary>
    Information,
    /// <summary>Finding requiring user review.</summary>
    Warning,
    /// <summary>Finding that prevented conversion.</summary>
    Error
}

/// <summary>Describes one structured PowerPoint conversion finding.</summary>
public sealed class PowerPointConversionDiagnostic {
    internal PowerPointConversionDiagnostic(
        string code,
        PowerPointConversionDiagnosticCategory category,
        PowerPointConversionDiagnosticSeverity severity,
        string message,
        OfficeCompatibilityState compatibilityState,
        OfficeCompatibilityImpact compatibilityImpact,
        bool representsDataLoss,
        string? sourceLocation = null,
        string? fallbackArtifact = null) {
        Code = code;
        Category = category;
        Severity = severity;
        Message = message;
        CompatibilityState = compatibilityState;
        CompatibilityImpact = compatibilityImpact;
        RepresentsDataLoss = representsDataLoss;
        SourceLocation = sourceLocation;
        FallbackArtifact = fallbackArtifact;
    }

    /// <summary>Gets the stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Gets the diagnostic category.</summary>
    public PowerPointConversionDiagnosticCategory Category { get; }

    /// <summary>Gets the diagnostic severity.</summary>
    public PowerPointConversionDiagnosticSeverity Severity { get; }

    /// <summary>Gets the human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Gets how the source feature is represented in the target.</summary>
    public OfficeCompatibilityState CompatibilityState { get; }

    /// <summary>Gets the affected fidelity dimensions.</summary>
    public OfficeCompatibilityImpact CompatibilityImpact { get; }

    /// <summary>Gets whether the finding represents source fidelity loss.</summary>
    public bool RepresentsDataLoss { get; }

    /// <summary>Gets the related source slide, shape, stream offset, or package part.</summary>
    public string? SourceLocation { get; }

    /// <summary>Gets the generated fallback artifact, when one exists.</summary>
    public string? FallbackArtifact { get; }
}

/// <summary>Represents the destination artifact and report produced by a PowerPoint conversion.</summary>
public sealed class PowerPointPresentationConversionResult {
    internal PowerPointPresentationConversionResult(
        string sourcePath,
        string destinationPath,
        PowerPointFileFormat sourceFormat,
        PowerPointFileFormat destinationFormat,
        OfficeFormatDescriptor sourceDescriptor,
        OfficeFormatDescriptor destinationDescriptor,
        OfficeCompatibilityMode compatibilityMode,
        IReadOnlyList<PowerPointConversionDiagnostic> diagnostics,
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
public sealed class PowerPointPresentationConversionReport {
    internal PowerPointPresentationConversionReport(
        string sourcePath,
        string destinationPath,
        PowerPointFileFormat sourceFormat,
        PowerPointFileFormat destinationFormat,
        OfficeFormatDescriptor sourceDescriptor,
        OfficeFormatDescriptor destinationDescriptor,
        OfficeCompatibilityMode compatibilityMode,
        IReadOnlyList<PowerPointConversionDiagnostic> diagnostics,
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
    public IReadOnlyList<PowerPointConversionDiagnostic> Diagnostics { get; }

    /// <summary>Gets the shared feature-level fidelity assessment.</summary>
    public OfficeCompatibilityReport Compatibility { get; }

    /// <summary>Gets whether the conversion reports fidelity loss.</summary>
    public bool HasLoss => Compatibility.HasLoss;

    /// <summary>Gets whether an existing destination was replaced.</summary>
    public bool ReplacedExistingFile { get; }

    /// <summary>Throws when the conversion reports loss or a blocked feature.</summary>
    public void RequireNoLoss() => Compatibility.RequireNoLoss();

    private static OfficeCompatibilityFinding ToCompatibilityFinding(PowerPointConversionDiagnostic diagnostic) => new(
        diagnostic.Code,
        diagnostic.Category.ToString(),
        diagnostic.Message,
        diagnostic.CompatibilityState,
        diagnostic.Severity switch {
            PowerPointConversionDiagnosticSeverity.Warning => OfficeCompatibilitySeverity.Warning,
            PowerPointConversionDiagnosticSeverity.Error => OfficeCompatibilitySeverity.Error,
            _ => OfficeCompatibilitySeverity.Information
        },
        diagnostic.CompatibilityImpact,
        diagnostic.RepresentsDataLoss,
        diagnostic.SourceLocation,
        diagnostic.FallbackArtifact);
}

/// <summary>Identifies why a PowerPoint conversion was rejected.</summary>
public enum PowerPointPresentationConversionFailureReason {
    /// <summary>The source already uses the requested concrete format.</summary>
    SameFormat,
    /// <summary>The destination exists and replacement was not allowed.</summary>
    DestinationExists,
    /// <summary>Known fidelity loss was blocked by policy.</summary>
    DataLossBlocked,
    /// <summary>The destination writer cannot represent source content.</summary>
    DestinationFeatureUnsupported
}

/// <summary>Raised when a validated PowerPoint conversion cannot be completed safely.</summary>
public sealed class PowerPointPresentationConversionException : InvalidOperationException {
    internal PowerPointPresentationConversionException(
        PowerPointPresentationConversionFailureReason reason,
        PowerPointPresentationConversionResult result,
        string message,
        Exception? innerException = null)
        : base(message, innerException) {
        Reason = reason;
        Result = result;
    }

    /// <summary>Gets the structured failure reason.</summary>
    public PowerPointPresentationConversionFailureReason Reason { get; }

    /// <summary>Gets the assessment available when conversion was rejected.</summary>
    public PowerPointPresentationConversionResult Result { get; }
}

/// <summary>Controls file-to-file PowerPoint conversion.</summary>
public sealed class PowerPointPresentationConversionOptions {
    /// <summary>Gets or sets how an existing destination is handled.</summary>
    public PowerPointConversionFileConflictPolicy FileConflictPolicy { get; set; } = PowerPointConversionFileConflictPolicy.FailIfExists;

    /// <summary>Gets or sets whether known conversion loss is blocked.</summary>
    public PowerPointConversionLossPolicy LossPolicy { get; set; } = PowerPointConversionLossPolicy.Block;

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
    public PowerPointSignatureMutationPolicy SignatureMutationPolicy { get; set; } =
        PowerPointSignatureMutationPolicy.BlockSave;

    /// <summary>Gets or sets optional source load settings.</summary>
    public PowerPointLoadOptions? LoadOptions { get; set; }

    /// <summary>Gets or sets optional destination save settings.</summary>
    public PowerPointSaveOptions? SaveOptions { get; set; }
}
