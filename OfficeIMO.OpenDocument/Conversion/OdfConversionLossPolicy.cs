namespace OfficeIMO.OpenDocument;

/// <summary>Policy applied after an explicit OpenDocument conversion has produced its evidence report.</summary>
public enum OdfConversionLossPolicy {
    /// <summary>Return the value and report; the caller decides how to handle loss.</summary>
    ReportOnly = 0,
    /// <summary>Throw when a feature was skipped or unsupported; documented approximations are accepted.</summary>
    ThrowOnSkippedOrUnsupported,
    /// <summary>Throw for every approximation, skip, or unsupported mapping.</summary>
    ThrowOnAnyLoss
}

/// <summary>Exception retaining the complete conversion report rejected by a loss policy.</summary>
public sealed class OdfConversionLossException : InvalidOperationException {
    /// <summary>Creates an exception for the rejected report.</summary>
    public OdfConversionLossException(OdfConversionReport report, string message) : base(message) {
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Complete feature-level conversion evidence.</summary>
    public OdfConversionReport Report { get; }
}
