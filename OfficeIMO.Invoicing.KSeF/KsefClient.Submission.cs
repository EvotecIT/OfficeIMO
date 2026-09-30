using System.Security.Cryptography;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.KSeF;

public sealed partial class KsefClient {
    /// <summary>Opens an online FA(3) session with a fresh AES-256 key/IV, wrapped using the active official RSA certificate.</summary>
    public async Task<KsefOnlineSession> OpenOnlineSessionAsync(KsefCredentials credentials, CancellationToken cancellationToken = default) {
        string token = Authorization(credentials); cancellationToken.ThrowIfCancellationRequested();
        using KsefEncryptionCertificate certificate = await CertificateAsync("SymmetricKeyEncryption", cancellationToken).ConfigureAwait(false);
        byte[] key = RandomNumberGenerator.GetBytes(32), iv = RandomNumberGenerator.GetBytes(16); bool transferred = false;
        try {
            byte[] encryptedKey = certificate.Key.Encrypt(key, RSAEncryptionPadding.OaepSHA256);
            byte[] payload = KsefProtocol.Json(writer => {
                writer.WriteStartObject("formCode"); writer.WriteString("systemCode", "FA (3)"); writer.WriteString("schemaVersion", "1-0E"); writer.WriteString("value", "FA"); writer.WriteEndObject();
                writer.WriteStartObject("encryption"); writer.WriteBase64String("encryptedSymmetricKey", encryptedKey); writer.WriteBase64String("initializationVector", iv); writer.WriteString("publicKeyId", certificate.PublicKeyId); writer.WriteEndObject();
            });
            KsefOnlineSession session = await JsonAsync(HttpMethod.Post, "sessions/online", payload, token, root => {
                string reference = KsefProtocol.Reference(KsefProtocol.Text(root, "referenceNumber", 36)); DateTimeOffset until = KsefProtocol.Instant(root, "validUntil");
                if (until <= _clock.GetUtcNow()) throw new InvalidDataException("Returned online session has expired.");
                return new KsefOnlineSession(_owner, credentials.Context, reference, until, key, iv);
            }, cancellationToken, true).ConfigureAwait(false);
            transferred = true; return session;
        } finally { if (!transferred) { CryptographicOperations.ZeroMemory(key); CryptographicOperations.ZeroMemory(iv); } }
    }
    /// <summary>Validates a bounded defensive FA(3) snapshot before network mutation, then submits its exact encrypted bytes and digests. A returned reference is not acceptance.</summary>
    public async Task<KsefSubmission> SubmitAsync(KsefCredentials credentials, KsefOnlineSession session, byte[] fa3Xml, Fa3SchemaValidator validator, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(session); ArgumentNullException.ThrowIfNull(fa3Xml); ArgumentNullException.ThrowIfNull(validator);
        Owned(session.Owner); string token = Authorization(credentials); SameContext(credentials.Context, session.Context);
        if (fa3Xml.Length == 0 || fa3Xml.Length > 16 * 1024 * 1024) throw new InvalidDataException("Invoice input must contain between one byte and 16 MiB.");
        byte[] snapshot = (byte[])fa3Xml.Clone();
        try {
            Fa3SchemaValidationResult validation = validator.Validate(snapshot, cancellationToken);
            if (!validation.IsValid) throw new InvalidDataException("Invoice does not pass the pinned FA(3) XSD; no submission occurred.");
            string hash = Convert.ToBase64String(SHA256.HashData(snapshot));
            await session.MutationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try {
                if (session.MutationStopped || _clock.GetUtcNow() >= session.ValidUntil) throw new InvalidOperationException("Session is closed, expired or has an ambiguous mutation; only reconciliation is permitted.");
                token = Authorization(credentials);
                byte[] encrypted = session.Encrypt(snapshot);
                byte[] payload = KsefProtocol.Json(writer => {
                    writer.WriteString("invoiceHash", hash); writer.WriteNumber("invoiceSize", snapshot.Length);
                    writer.WriteBase64String("encryptedInvoiceHash", SHA256.HashData(encrypted)); writer.WriteNumber("encryptedInvoiceSize", encrypted.Length);
                    writer.WriteBase64String("encryptedInvoiceContent", encrypted); writer.WriteBoolean("offlineMode", false);
                });
                try {
                    return await JsonAsync(HttpMethod.Post, "sessions/online/" + Uri.EscapeDataString(session.ReferenceNumber) + "/invoices", payload, token,
                        root => new KsefSubmission(_owner, credentials.Context, session.ReferenceNumber, KsefProtocol.Reference(KsefProtocol.Text(root, "referenceNumber", 36)), hash),
                        cancellationToken, true, session.ReferenceNumber, hash).ConfigureAwait(false);
                } catch (KsefMutationAmbiguousException) { session.MutationStopped = true; throw; }
            } finally { session.MutationGate.Release(); }
        } finally { CryptographicOperations.ZeroMemory(snapshot); }
    }
    /// <summary>Closes a session once. An ambiguous close stops further mutation; status remains readable by reference.</summary>
    public async Task CloseOnlineSessionAsync(KsefCredentials credentials, KsefOnlineSession session, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(session); Owned(session.Owner); string token = Authorization(credentials); SameContext(credentials.Context, session.Context);
        await session.MutationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (session.MutationStopped) throw new InvalidOperationException("Session mutation has already stopped; reconcile its status instead of replaying close.");
            token = Authorization(credentials);
            try {
                await SendAsync(HttpMethod.Post, "sessions/online/" + Uri.EscapeDataString(session.ReferenceNumber) + "/close", null, token, 65_536, cancellationToken, true, session.ReferenceNumber).ConfigureAwait(false);
                session.MutationStopped = true; session.Dispose();
            } catch (KsefMutationAmbiguousException) { session.MutationStopped = true; session.Dispose(); throw; }
        } finally { session.MutationGate.Release(); }
    }
    private static void SameContext(KsefContext left, KsefContext right) {
        if (left.Kind != right.Kind || left.Value != right.Value) throw new InvalidOperationException("KSeF contexts cannot be mixed.");
    }
}
