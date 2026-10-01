using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Invoicing.KSeF;

/// <summary>An owned secret buffer. Disposal clears owned bytes; the caller's original string and runtime HTTP header strings cannot be erased by this type.</summary>
public sealed class KsefSecret : IDisposable {
    private readonly object _gate = new();
    private byte[]? _bytes;
    /// <summary>Copies a nonempty secret without control characters, bounded to 32 KiB in UTF-8.</summary>
    public KsefSecret(string value) {
        ArgumentNullException.ThrowIfNull(value);
        if (value.Length == 0 || value.Length > 32_768 || value.Any(char.IsControl) || Encoding.UTF8.GetByteCount(value) > 32_768)
            throw new ArgumentException("Secret must be nonempty, contain no control characters and fit within 32 KiB.", nameof(value));
        _bytes = Encoding.UTF8.GetBytes(value);
    }
    internal string HeaderValue() { lock (_gate) return Encoding.UTF8.GetString(_bytes ?? throw new ObjectDisposedException(nameof(KsefSecret))); }
    internal void CheckOpen() { lock (_gate) { if (_bytes == null) throw new ObjectDisposedException(nameof(KsefSecret)); } }
    internal byte[] EncryptToken(RSA key, long timestamp) {
        lock (_gate) {
            byte[] bytes = _bytes ?? throw new ObjectDisposedException(nameof(KsefSecret));
            byte[] suffix = Encoding.ASCII.GetBytes("|" + timestamp.ToString(System.Globalization.CultureInfo.InvariantCulture));
            byte[] input = new byte[bytes.Length + suffix.Length];
            try { bytes.CopyTo(input, 0); suffix.CopyTo(input, bytes.Length); return key.Encrypt(input, RSAEncryptionPadding.OaepSHA256); }
            finally { CryptographicOperations.ZeroMemory(input); CryptographicOperations.ZeroMemory(suffix); }
        }
    }
    /// <summary>Never formats the secret.</summary>
    public override string ToString() => "[redacted]";
    /// <summary>Clears the owned buffer. Already dispatched HTTP requests are unaffected.</summary>
    public void Dispose() { lock (_gate) { if (_bytes != null) CryptographicOperations.ZeroMemory(_bytes); _bytes = null; } }
}
