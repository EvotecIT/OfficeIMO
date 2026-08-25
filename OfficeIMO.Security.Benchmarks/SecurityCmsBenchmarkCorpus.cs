using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security.Benchmarks;

internal static class SecurityCmsBenchmarkCorpus {
    internal static readonly string[] Scales = { "Small", "Normal", "Large" };

    internal static SecurityCmsBenchmarkFixture Create(string scale) {
        int contentBytes = scale switch {
            "Small" => 1_024,
            "Normal" => 65_536,
            "Large" => 1_048_576,
            _ => throw new ArgumentOutOfRangeException(nameof(scale), scale, "Unknown CMS benchmark scale.")
        };

        var content = new byte[contentBytes];
        for (int index = 0; index < content.Length; index++) {
            content[index] = (byte)((index * 31L + index / 251) % 251);
        }

        RSA key = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO CMS Benchmark",
            key,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, false));
        X509Certificate2 certificate = request.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddDays(-1),
            DateTimeOffset.UtcNow.AddDays(7));
        return new SecurityCmsBenchmarkFixture(scale, content, key, certificate);
    }
}

internal sealed class SecurityCmsBenchmarkFixture : IDisposable {
    internal SecurityCmsBenchmarkFixture(string scale, byte[] content, RSA key, X509Certificate2 certificate) {
        Scale = scale;
        Content = content;
        Key = key;
        Certificate = certificate;
    }

    internal string Scale { get; }
    internal byte[] Content { get; }
    internal RSA Key { get; }
    internal X509Certificate2 Certificate { get; }

    public void Dispose() {
        Certificate.Dispose();
        Key.Dispose();
    }
}
