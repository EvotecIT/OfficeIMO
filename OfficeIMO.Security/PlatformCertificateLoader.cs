using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

internal static class PlatformCertificateLoader {
    internal static X509Certificate2 Load(byte[] encodedCertificate) {
#if NET9_0_OR_GREATER
        return X509CertificateLoader.LoadCertificate(encodedCertificate);
#else
        return new X509Certificate2(encodedCertificate);
#endif
    }
}
