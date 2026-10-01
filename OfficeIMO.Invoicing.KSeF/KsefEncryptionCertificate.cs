using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;

namespace OfficeIMO.Invoicing.KSeF;

internal sealed class KsefEncryptionCertificate : IDisposable {
    private KsefEncryptionCertificate(RSA rsa, string id) { Key = rsa; PublicKeyId = id; }
    internal RSA Key { get; }
    internal string PublicKeyId { get; }
    internal static KsefEncryptionCertificate Select(JsonElement certificates, string usage, DateTimeOffset now) {
        if (certificates.ValueKind != JsonValueKind.Array || certificates.GetArrayLength() > 64) throw new InvalidDataException("KSeF certificate list exceeds the supported bound.");
        foreach (JsonElement entry in certificates.EnumerateArray()) {
            JsonElement uses = entry.GetProperty("usage");
            if (uses.ValueKind != JsonValueKind.Array || uses.GetArrayLength() > 8) throw new InvalidDataException("Invalid KSeF key usage list.");
            if (!uses.EnumerateArray().Any(value => value.ValueKind == JsonValueKind.String && value.GetString() == usage)) continue;
            DateTimeOffset from = KsefProtocol.Instant(entry, "validFrom"), until = KsefProtocol.Instant(entry, "validTo");
            if (now < from || now >= until) continue;
            string encoded = KsefProtocol.Text(entry, "certificate", 43_692);
            byte[] bytes = Convert.FromBase64String(encoded);
            if (bytes.Length == 0 || bytes.Length > 32_768 || Convert.ToBase64String(SHA256.HashData(bytes)) != KsefProtocol.Hash(KsefProtocol.Text(entry, "certificateId", 44))) throw new InvalidDataException("KSeF certificate digest mismatch.");
#if NET9_0_OR_GREATER
            using X509Certificate2 certificate = X509CertificateLoader.LoadCertificate(bytes);
#else
            using var certificate = new X509Certificate2(bytes);
#endif
            if (now < certificate.NotBefore.ToUniversalTime() || now >= certificate.NotAfter.ToUniversalTime()) continue;
            RSA? rsa = certificate.GetRSAPublicKey();
            if (rsa == null) continue;
            string id = KsefProtocol.Hash(KsefProtocol.Text(entry, "publicKeyId", 44));
            if (rsa.KeySize < 2048 || Convert.ToBase64String(SHA256.HashData(rsa.ExportSubjectPublicKeyInfo())) != id) { rsa.Dispose(); throw new InvalidDataException("KSeF RSA key or SPKI digest is invalid."); }
            return new KsefEncryptionCertificate(rsa, id);
        }
        throw new InvalidDataException("No active RSA certificate supports the required KSeF encryption usage.");
    }
    public void Dispose() => Key.Dispose();
}
