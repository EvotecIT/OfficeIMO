using System.Text.Json;

namespace OfficeIMO.Invoicing.KSeF;

public sealed partial class KsefClient {
    /// <summary>Restores a status/receipt reference after a process restart using newly authenticated credentials. This neither reopens a session nor resubmits XML.</summary>
    public KsefSubmission ResumeSubmission(KsefCredentials credentials, string sessionReference, string invoiceReference, string invoiceHash) {
        Authorization(credentials);
        return new KsefSubmission(_owner, credentials.Context, KsefProtocol.Reference(sessionReference), KsefProtocol.Reference(invoiceReference), KsefProtocol.Hash(invoiceHash));
    }
    /// <summary>Reads a session status by reference, including after close or an ambiguous mutation.</summary>
    public Task<KsefStatus> GetSessionStatusAsync(KsefCredentials credentials, string sessionReference, CancellationToken cancellationToken = default) =>
        JsonAsync(HttpMethod.Get, "sessions/" + Uri.EscapeDataString(KsefProtocol.Reference(sessionReference)), null, Authorization(credentials), KsefProtocol.Status, cancellationToken);
    /// <summary>Reads invoice status and rejects a response bound to a different reference or plaintext hash.</summary>
    public Task<KsefInvoiceStatus> GetInvoiceStatusAsync(KsefCredentials credentials, KsefSubmission submission, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(submission); Owned(submission.Owner); string token = Authorization(credentials); SameContext(credentials.Context, submission.Context);
        return JsonAsync(HttpMethod.Get, InvoicePath(submission), null, token, root => {
            if (KsefProtocol.Text(root, "referenceNumber", 36) != submission.ReferenceNumber || KsefProtocol.Hash(KsefProtocol.Text(root, "invoiceHash", 44)) != submission.InvoiceHash)
                throw new InvalidDataException("Invoice status does not match the expected reference and exact plaintext hash.");
            return InvoiceStatus(root);
        }, cancellationToken);
    }
    /// <summary>Lists one bounded page for explicit ambiguous-submission reconciliation. It never resends an invoice and never follows a remote download URL.</summary>
    public Task<KsefInvoicePage> GetSessionInvoicesAsync(KsefCredentials credentials, string sessionReference, int pageSize = 100, string? continuationToken = null, CancellationToken cancellationToken = default) {
        if (pageSize < 10 || pageSize > 1000) throw new ArgumentOutOfRangeException(nameof(pageSize));
        string reference = KsefProtocol.Reference(sessionReference);
        return JsonAsync(HttpMethod.Get, "sessions/" + Uri.EscapeDataString(reference) + "/invoices?pageSize=" + pageSize.ToString(System.Globalization.CultureInfo.InvariantCulture), null, Authorization(credentials), root => {
            JsonElement invoices = root.GetProperty("invoices");
            if (invoices.ValueKind != JsonValueKind.Array || invoices.GetArrayLength() > pageSize) throw new InvalidDataException("Invoice page exceeds the requested bound.");
            var rows = new List<KsefSessionInvoice>();
            foreach (JsonElement invoice in invoices.EnumerateArray()) {
                var submission = new KsefSubmission(_owner, credentials.Context, reference, KsefProtocol.Reference(KsefProtocol.Text(invoice, "referenceNumber", 36)), KsefProtocol.Hash(KsefProtocol.Text(invoice, "invoiceHash", 44)));
                rows.Add(new KsefSessionInvoice(submission, InvoiceStatus(invoice)));
            }
            string? continuation = root.TryGetProperty("continuationToken", out JsonElement next) && next.ValueKind != JsonValueKind.Null ? next.GetString() : null;
            if (continuation != null && (continuation.Length > 8192 || continuation.Any(char.IsControl))) throw new InvalidDataException("Invalid invoice continuation token.");
            return new KsefInvoicePage(rows.AsReadOnly(), continuation);
        }, cancellationToken, continuation: continuationToken);
    }
    /// <summary>Reads UPO through the direct authenticated official route after accepted status, with separate pinned schema and exact context/session/document binding results.</summary>
    public async Task<KsefReceipt> GetInvoiceReceiptAsync(KsefCredentials credentials, KsefSubmission submission, CancellationToken cancellationToken = default) {
        KsefInvoiceStatus status = await GetInvoiceStatusAsync(credentials, submission, cancellationToken).ConfigureAwait(false);
        if (!status.IsAccepted) throw new InvalidOperationException("Invoice acceptance is not established; a UPO is not claimed.");
        byte[] bytes = await SendAsync(HttpMethod.Get, InvoicePath(submission) + "/upo", null, Authorization(credentials), 2 * 1024 * 1024, cancellationToken).ConfigureAwait(false);
        return KsefReceipt.Read(bytes, credentials.Context, submission.SessionReference, status.KsefNumber!, submission.InvoiceHash, true, cancellationToken);
    }
    private static string InvoicePath(KsefSubmission submission) => "sessions/" + Uri.EscapeDataString(submission.SessionReference) + "/invoices/" + Uri.EscapeDataString(submission.ReferenceNumber);
    private static KsefInvoiceStatus InvoiceStatus(JsonElement root) {
        string? ksefNumber = root.TryGetProperty("ksefNumber", out JsonElement number) && number.ValueKind != JsonValueKind.Null ? KsefProtocol.InvoiceNumber(number.GetString()!) : null;
        KsefStatus status = KsefProtocol.Status(root);
        if (status.IsSuccessful && ksefNumber == null) throw new InvalidDataException("Successful invoice status lacks a KSeF number.");
        return new KsefInvoiceStatus(status, ksefNumber, KsefProtocol.Instant(root, "invoicingDate"));
    }
}
