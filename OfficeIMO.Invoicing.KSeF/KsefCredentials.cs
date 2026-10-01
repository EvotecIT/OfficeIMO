namespace OfficeIMO.Invoicing.KSeF;

/// <summary>One authentication operation and its temporary token, bound to its originating client.</summary>
public sealed class KsefAuthentication : IDisposable {
    private int _redeemed;
    internal KsefAuthentication(Guid owner, KsefContext context, string reference, KsefSecret token, DateTimeOffset validUntil) { Owner = owner; Context = context; ReferenceNumber = reference; Token = token; ValidUntil = validUntil; }
    internal Guid Owner { get; }
    internal KsefSecret Token { get; }
    internal KsefContext Context { get; }
    internal void ReserveRedemption() { if (Interlocked.CompareExchange(ref _redeemed, 1, 0) != 0) throw new InvalidOperationException("Redemption has already been attempted; automatic replay is prohibited."); }
    /// <summary>Authentication operation reference.</summary>
    public string ReferenceNumber { get; }
    /// <summary>Temporary authentication-token expiration.</summary>
    public DateTimeOffset ValidUntil { get; }
    /// <summary>Clears the owned temporary token.</summary>
    public void Dispose() => Token.Dispose();
}
/// <summary>Access and refresh tokens without publicly readable secret properties. Credentials are bound to one client and context.</summary>
public sealed class KsefCredentials : IDisposable {
    private readonly object _gate = new();
    private KsefSecret? _access;
    private KsefSecret? _refresh;
    internal readonly SemaphoreSlim RefreshGate = new(1, 1);
    internal KsefCredentials(Guid owner, KsefContext context, KsefSecret access, DateTimeOffset accessUntil, KsefSecret refresh, DateTimeOffset refreshUntil) {
        Owner = owner; Context = context; _access = access; AccessValidUntil = accessUntil; _refresh = refresh; RefreshValidUntil = refreshUntil;
    }
    internal Guid Owner { get; }
    /// <summary>Authenticated context; its value is an identifier, not a token.</summary>
    public KsefContext Context { get; }
    /// <summary>Expiration of the current access token.</summary>
    public DateTimeOffset AccessValidUntil { get; private set; }
    /// <summary>Expiration of the refresh token.</summary>
    public DateTimeOffset RefreshValidUntil { get; }
    internal string Header(DateTimeOffset now, bool refresh = false) {
        lock (_gate) {
            if (now >= (refresh ? RefreshValidUntil : AccessValidUntil)) throw new InvalidOperationException("KSeF credential has expired.");
            return (refresh ? _refresh : _access)?.HeaderValue() ?? throw new ObjectDisposedException(nameof(KsefCredentials));
        }
    }
    internal void ReplaceAccess(KsefSecret token, DateTimeOffset until) {
        lock (_gate) {
            if (_refresh == null) { token.Dispose(); throw new ObjectDisposedException(nameof(KsefCredentials)); }
            _access?.Dispose(); _access = token; AccessValidUntil = until;
        }
    }
    /// <summary>Clears both owned token buffers.</summary>
    public void Dispose() { lock (_gate) { _access?.Dispose(); _refresh?.Dispose(); _access = null; _refresh = null; } }
}
