namespace OfficeIMO.Email;

/// <summary>Controls whether a conversion may continue when message semantics cannot be preserved.</summary>
public enum EmailConversionLossPolicy {
    /// <summary>Stop before producing output when a known semantic loss would occur.</summary>
    Block = 0,
    /// <summary>Produce output and report each known semantic loss as a warning.</summary>
    Warn = 1,
    /// <summary>Produce output and report each accepted semantic loss as information.</summary>
    Allow = 2
}
