namespace OfficeIMO.Email;

/// <summary>How processing continued after the condition represented by a diagnostic.</summary>
public enum EmailDiagnosticDisposition {
    /// <summary>The condition was observed without changing the processing path.</summary>
    Observed = 0,
    /// <summary>The reader recovered the affected structure and continued.</summary>
    Recovered = 1,
    /// <summary>The affected item or value was skipped and processing continued.</summary>
    Skipped = 2,
    /// <summary>Processing stopped because continuing would be unsafe or misleading.</summary>
    Stopped = 3
}
/// <summary>Data-loss consequence associated with a diagnostic.</summary>
public enum EmailDataLossRisk {
    /// <summary>No user data was lost by the reported condition.</summary>
    None = 0,
    /// <summary>Some content may be unavailable or may not round-trip.</summary>
    Possible = 1,
    /// <summary>The operation intentionally or unavoidably omitted content.</summary>
    Confirmed = 2,
    /// <summary>The parser could not determine the loss consequence safely.</summary>
    Unknown = 3
}
