namespace OfficeIMO.Rtf;

/// <summary>Severity of a cross-format RTF conversion diagnostic.</summary>
public enum RtfConversionSeverity {
    /// <summary>Informational conversion detail.</summary>
    Information,

    /// <summary>Conversion completed with a noteworthy condition or fidelity loss.</summary>
    Warning,

    /// <summary>Conversion could not preserve a required contract.</summary>
    Error
}
