using System.Security.Cryptography;

namespace OfficeIMO.Invoicing.KSeF;

/// <summary>An online session owning its AES-256 key and IV. Submission and close are serialized; ambiguous mutations stop further mutation of this handle.</summary>
public sealed class KsefOnlineSession : IDisposable {
    private readonly object _cryptoGate = new();
    private byte[]? _key;
    private byte[]? _iv;
    internal readonly SemaphoreSlim MutationGate = new(1, 1);
    internal bool MutationStopped;
    internal KsefOnlineSession(Guid owner, KsefContext context, string reference, DateTimeOffset until, byte[] key, byte[] iv) {
        Owner = owner; Context = context; ReferenceNumber = reference; ValidUntil = until; _key = key; _iv = iv;
    }
    internal Guid Owner { get; }
    internal KsefContext Context { get; }
    /// <summary>Official session reference.</summary>
    public string ReferenceNumber { get; }
    /// <summary>Server-declared session expiration.</summary>
    public DateTimeOffset ValidUntil { get; }
    internal byte[] Encrypt(byte[] input) {
        lock (_cryptoGate) {
            using Aes aes = Aes.Create(); aes.Key = _key ?? throw new ObjectDisposedException(nameof(KsefOnlineSession));
            return aes.EncryptCbc(input, _iv ?? throw new ObjectDisposedException(nameof(KsefOnlineSession)), PaddingMode.PKCS7);
        }
    }
    /// <summary>Clears owned key material. An already dispatched operation may still complete; this does not close the remote session.</summary>
    public void Dispose() {
        lock (_cryptoGate) { if (_key != null) CryptographicOperations.ZeroMemory(_key); if (_iv != null) CryptographicOperations.ZeroMemory(_iv); _key = null; _iv = null; }
    }
}
