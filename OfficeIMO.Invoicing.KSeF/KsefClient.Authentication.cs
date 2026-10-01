using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Invoicing.KSeF;

public sealed partial class KsefClient {
    /// <summary>Checks the current public encryption certificates without credentials, returning the verified active RSA SPKI selector.</summary>
    public async Task<string> CheckEncryptionKeyAsync(CancellationToken cancellationToken = default) {
        using KsefEncryptionCertificate certificate = await CertificateAsync("SymmetricKeyEncryption", cancellationToken).ConfigureAwait(false); return certificate.PublicKeyId;
    }
    private Task<KsefEncryptionCertificate> CertificateAsync(string usage, CancellationToken cancellationToken) =>
        JsonAsync(HttpMethod.Get, "security/public-key-certificates", null, null, root => KsefEncryptionCertificate.Select(root, usage, _clock.GetUtcNow()), cancellationToken);

    /// <summary>Starts token authentication using the server challenge timestamp and active token-encryption RSA certificate. Success is asynchronous.</summary>
    public async Task<KsefAuthentication> BeginTokenAuthenticationAsync(KsefContext context, KsefSecret ksefToken, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(context); ArgumentNullException.ThrowIfNull(ksefToken); ksefToken.CheckOpen(); cancellationToken.ThrowIfCancellationRequested();
        using KsefEncryptionCertificate certificate = await CertificateAsync("KsefTokenEncryption", cancellationToken).ConfigureAwait(false);
        var challenge = await JsonAsync(HttpMethod.Post, "auth/challenge", KsefProtocol.Json(_ => { }), null, root => {
            string value = KsefProtocol.Challenge(KsefProtocol.Text(root, "challenge", 36));
            long timestamp = root.GetProperty("timestampMs").GetInt64(); DateTimeOffset instant = KsefProtocol.Instant(root, "timestamp");
            if (instant.ToUnixTimeMilliseconds() != timestamp || instant < _clock.GetUtcNow().AddMinutes(-10) || instant > _clock.GetUtcNow().AddSeconds(30)) throw new InvalidDataException("Authentication challenge timestamp is inconsistent or expired.");
            return (Value: value, Timestamp: timestamp);
        }, cancellationToken).ConfigureAwait(false);
        byte[] encrypted = ksefToken.EncryptToken(certificate.Key, challenge.Timestamp);
        byte[] payload = KsefProtocol.Json(writer => {
            writer.WriteString("challenge", challenge.Value); writer.WriteStartObject("contextIdentifier"); writer.WriteString("type", context.Kind.ToString()); writer.WriteString("value", context.Value); writer.WriteEndObject();
            writer.WriteBase64String("encryptedToken", encrypted); writer.WriteString("publicKeyId", certificate.PublicKeyId);
        });
        return await JsonAsync(HttpMethod.Post, "auth/ksef-token", payload, null, root => {
            string reference = KsefProtocol.Reference(KsefProtocol.Text(root, "referenceNumber", 36));
            (KsefSecret token, DateTimeOffset until) = Token(root.GetProperty("authenticationToken"));
            return new KsefAuthentication(_owner, context, reference, token, until);
        }, cancellationToken, mutation: true).ConfigureAwait(false);
    }
    /// <summary>Reads authentication state without redeeming tokens or replaying authentication.</summary>
    public Task<KsefStatus> GetAuthenticationStatusAsync(KsefAuthentication authentication, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(authentication); Owned(authentication.Owner);
        if (_clock.GetUtcNow() >= authentication.ValidUntil) throw new InvalidOperationException("Authentication token has expired.");
        return JsonAsync(HttpMethod.Get, "auth/" + Uri.EscapeDataString(authentication.ReferenceNumber), null, authentication.Token.HeaderValue(), KsefProtocol.Status, cancellationToken);
    }
    /// <summary>Redeems a successful authentication once. Lost/malformed responses leave this handle consumed because server-side redemption may already have happened.</summary>
    public async Task<KsefCredentials> RedeemAsync(KsefAuthentication authentication, CancellationToken cancellationToken = default) {
        KsefStatus status = await GetAuthenticationStatusAsync(authentication, cancellationToken).ConfigureAwait(false);
        if (!status.IsSuccessful) throw new InvalidOperationException("Authentication has not completed successfully.");
        return await JsonAsync(HttpMethod.Post, "auth/token/redeem", null, authentication.Token.HeaderValue(), root => {
            (KsefSecret access, DateTimeOffset accessUntil) = Token(root.GetProperty("accessToken"));
            try {
                (KsefSecret refresh, DateTimeOffset refreshUntil) = Token(root.GetProperty("refreshToken"));
                return new KsefCredentials(_owner, authentication.Context, access, accessUntil, refresh, refreshUntil);
            } catch { access.Dispose(); throw; }
        }, cancellationToken, true, authentication.ReferenceNumber, beforeDispatch: authentication.ReserveRedemption).ConfigureAwait(false);
    }
    /// <summary>Refreshes access credentials using the current refresh token. No automatic retry occurs.</summary>
    public async Task RefreshAsync(KsefCredentials credentials, CancellationToken cancellationToken = default) {
        Authorization(credentials, true);
        await credentials.RefreshGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            string token = Authorization(credentials, true);
            var access = await JsonAsync(HttpMethod.Post, "auth/token/refresh", null, token, root => Token(root.GetProperty("accessToken")), cancellationToken, true).ConfigureAwait(false);
            credentials.ReplaceAccess(access.Token, access.Until);
        } finally { credentials.RefreshGate.Release(); }
    }
    private (KsefSecret Token, DateTimeOffset Until) Token(JsonElement element) {
        DateTimeOffset until = KsefProtocol.Instant(element, "validUntil");
        if (until <= _clock.GetUtcNow()) throw new InvalidDataException("Returned KSeF token has expired.");
        return (new KsefSecret(KsefProtocol.Text(element, "token", 32_768)), until);
    }
}
