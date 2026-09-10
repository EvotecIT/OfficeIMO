namespace OfficeIMO.Invoicing.Validation;

/// <summary>Explicit pinned business-rule releases. There is no moving "latest" selection.</summary>
public enum InvoiceRulesRelease {
    /// <summary>EN 16931 validation artifacts 1.3.16, CII D16B or UBL 2.1.</summary>
    En16931_1_3_16,
    /// <summary>XRechnung 3.0.2 configuration 2026-08-31, including its severity overrides.</summary>
    XRechnung_3_0_2_2026_08_31,
    /// <summary>Peppol BIS Billing 3.0.21 (May 2026) with EN 16931 1.3.16, UBL only.</summary>
    PeppolBis_3_0_21
}

/// <summary>Outcome of one validation stage.</summary>
public enum InvoiceValidationStatus {
    /// <summary>The stage did not run.</summary>
    NotRun,
    /// <summary>The stage ran and found no errors.</summary>
    Passed,
    /// <summary>The stage ran and found invalid invoice content.</summary>
    Invalid,
    /// <summary>The validation engine could not complete the stage.</summary>
    Failed
}
