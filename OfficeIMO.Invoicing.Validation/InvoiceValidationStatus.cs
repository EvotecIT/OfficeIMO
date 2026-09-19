namespace OfficeIMO.Invoicing.Validation;

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
