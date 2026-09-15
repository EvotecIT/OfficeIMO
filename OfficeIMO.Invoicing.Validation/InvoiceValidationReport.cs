namespace OfficeIMO.Invoicing.Validation;

/// <summary>Validation evidence bound to exact invoice bytes and an explicit authority release.</summary>
public sealed class InvoiceValidationReport {
    internal InvoiceValidationReport(string sha256, int length, InvoiceSpecificationRelease release, InvoiceValidationStatus schema, InvoiceValidationStatus rules,
        IReadOnlyList<InvoiceDiagnostic> diagnostics, string? runner) {
        Sha256 = sha256; Length = length; Release = release; SchemaStatus = schema; BusinessRulesStatus = rules; Diagnostics = diagnostics; Runner = runner;
    }
    /// <summary>SHA-256 of the exact invoice bytes validated.</summary>
    public string Sha256 { get; }
    /// <summary>Length of the exact invoice bytes validated.</summary>
    public int Length { get; }
    /// <summary>Pinned rule release used by this run.</summary>
    public InvoiceSpecificationRelease Release { get; }
    /// <summary>Native XML Schema validation outcome.</summary>
    public InvoiceValidationStatus SchemaStatus { get; }
    /// <summary>Schematron validation outcome.</summary>
    public InvoiceValidationStatus BusinessRulesStatus { get; }
    /// <summary>Engine identity after an invoice-rule process started, including failed executions. Null for startup failure or compiler-only execution.</summary>
    public string? Runner { get; }
    /// <summary>All schema, engine and rule diagnostics.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
    /// <summary>True only when both schema and business rules ran successfully and found no errors.</summary>
    public bool IsValid => SchemaStatus == InvoiceValidationStatus.Passed && BusinessRulesStatus == InvoiceValidationStatus.Passed;
    /// <summary>SHA-256 of the authority bundle that supplied this release's schema and primary rules.</summary>
    public string AuthorityBundleSha256 => Release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2
        ? InvoiceRuleBundle.FacturXSha256
        : InvoiceRuleBundle.XRechnungSha256;
    /// <summary>Immutable source used for the authority bundle.</summary>
    public string AuthoritySource => Release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2
        ? InvoiceRuleBundle.FacturXDownloadUrl
        : InvoiceRuleBundle.XRechnungDownloadUrl;
    /// <summary>Source commit when the selected authority bundle is pinned directly to a repository commit.</summary>
    public string? AuthoritySourceCommit => Release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2
        ? InvoiceRuleBundle.FacturXSourceCommit
        : null;
    /// <summary>Additional Peppol source hash when the selected release needs it.</summary>
    public string? PeppolSourceSha256 => Release == InvoiceSpecificationRelease.PeppolBis_3_0_21 ? InvoiceRuleBundle.PeppolSha256 : null;
}
