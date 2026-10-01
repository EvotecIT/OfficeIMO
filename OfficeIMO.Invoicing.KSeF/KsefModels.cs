using System.Net;

namespace OfficeIMO.Invoicing.KSeF;

/// <summary>Explicit official API destination. TEST is the client default.</summary>
public enum KsefEnvironment {
    /// <summary>Credential and invoice test environment.</summary>
    Test,
    /// <summary>Preproduction demonstration environment.</summary>
    Demo,
    /// <summary>Production environment; selecting it is an explicit caller decision.</summary>
    Production
}
/// <summary>Official authentication context identifier kinds.</summary>
public enum KsefContextKind {
    /// <summary>Polish tax identifier.</summary>
    Nip,
    /// <summary>Internal identifier.</summary>
    InternalId,
    /// <summary>Polish NIP plus EU VAT identifier.</summary>
    NipVatUe,
    /// <summary>Registered Peppol provider identifier.</summary>
    PeppolId
}
/// <summary>An immutable context validated against the pinned authentication schema.</summary>
public sealed class KsefContext {
    /// <summary>Creates a context without verifying its registration or permissions.</summary>
    public KsefContext(KsefContextKind kind, string value) {
        if (!Enum.IsDefined(kind) || string.IsNullOrEmpty(value) || value.Length > 64) throw new ArgumentException("Supply a supported context kind and bounded identifier.");
        Kind = kind; Value = value; KsefProtocolSchemas.CheckContext(this);
    }
    /// <summary>Official identifier kind.</summary>
    public KsefContextKind Kind { get; }
    /// <summary>Identifier value; this is not an authentication secret.</summary>
    public string Value { get; }
}
/// <summary>A rejected HTTP operation without echoing server bodies, request secrets or invoice content.</summary>
public sealed class KsefApiException : Exception {
    internal KsefApiException(HttpStatusCode status, TimeSpan? retryAfter = null) : base("KSeF returned HTTP " + (int)status + ".") { StatusCode = status; RetryAfter = retryAfter; }
    /// <summary>Returned HTTP status.</summary>
    public HttpStatusCode StatusCode { get; }
    /// <summary>Server-supplied retry delay, if present. The client does not automatically retry.</summary>
    public TimeSpan? RetryAfter { get; }
}
/// <summary>A mutation whose outcome cannot be inferred from the response. Do not replay automatically; reconcile through status or start a fresh authentication.</summary>
public sealed class KsefMutationAmbiguousException : Exception {
    internal KsefMutationAmbiguousException(string operation, string? reference, string? invoiceHash) : base("KSeF " + operation + " has an unknown remote outcome; automatic replay is disabled.") {
        Operation = operation; ReferenceNumber = reference; InvoiceHash = invoiceHash;
    }
    /// <summary>Protocol operation without secrets.</summary>
    public string Operation { get; }
    /// <summary>Known session/authentication reference, if available.</summary>
    public string? ReferenceNumber { get; }
    /// <summary>Exact original invoice SHA-256 in Base64 when applicable.</summary>
    public string? InvoiceHash { get; }
}
/// <summary>Typed operation status. Descriptions and details remain remote text and are not included in exceptions.</summary>
public sealed record KsefStatus(int Code, string Description) {
    /// <summary>Whether the official success code is present.</summary>
    public bool IsSuccessful => Code == 200;
    /// <summary>Whether known asynchronous processing codes are present.</summary>
    public bool IsPending => Code is 100 or 150;
}
/// <summary>Immutable identity of a submitted document. Submission reference alone does not establish acceptance.</summary>
public sealed class KsefSubmission {
    internal KsefSubmission(Guid owner, KsefContext context, string session, string reference, string hash) { Owner = owner; Context = context; SessionReference = session; ReferenceNumber = reference; InvoiceHash = hash; }
    internal Guid Owner { get; }
    /// <summary>Expected authenticated context for status and receipt binding.</summary>
    public KsefContext Context { get; }
    /// <summary>Online session reference.</summary>
    public string SessionReference { get; }
    /// <summary>Invoice operation reference.</summary>
    public string ReferenceNumber { get; }
    /// <summary>SHA-256 of exact submitted plaintext bytes, Base64 encoded.</summary>
    public string InvoiceHash { get; }
}
/// <summary>Status tied to the exact submitted invoice hash and reference.</summary>
public sealed record KsefInvoiceStatus(KsefStatus Status, string? KsefNumber, DateTimeOffset InvoicingDate) {
    /// <summary>Successful status with an actual KSeF number; UPO qualification remains a separate step.</summary>
    public bool IsAccepted => Status.IsSuccessful && KsefNumber != null;
}
/// <summary>One remotely listed document with its exact plaintext hash and operation status.</summary>
public sealed record KsefSessionInvoice(KsefSubmission Submission, KsefInvoiceStatus Status);
/// <summary>A bounded session invoice page used to reconcile an ambiguous submission without replaying it.</summary>
public sealed record KsefInvoicePage(IReadOnlyList<KsefSessionInvoice> Invoices, string? ContinuationToken);
