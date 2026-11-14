using OfficeIMO.Security;

namespace OfficeIMO.Security.Tests;

public sealed class PublicKeySignatureVerifierTests {
    private static readonly IOfficePublicKeySignatureVerifier Verifier = OfficeSecurityProvider.Default;

    [Fact]
    public void Ed25519MatchesRfc8032TestVectorOne() {
        // RFC 8032 section 7.1, TEST 1 (empty message).
        byte[] publicKey = Convert.FromHexString("d75a980182b10ab7d54bfed3c964073a0ee172f3daa62325af021a68f707511a");
        byte[] signature = Convert.FromHexString(
            "e5564300c360ac729086e2cc806e828a84877f1eb8e5d974d873e06522490155" +
            "5fb8821590a33bacc61e39701cf9b46bd25bf5f0595bbe24655141438e7a100b");

        Assert.True(Verifier.Verify(OfficePublicKeySignatureAlgorithm.Ed25519, publicKey, Array.Empty<byte>(), signature).IsValid);
        Assert.Equal(OfficePublicKeySignatureStatus.Invalid,
            Verifier.Verify(OfficePublicKeySignatureAlgorithm.Ed25519, publicKey, new byte[] { 0 }, signature).Status);
    }

    [Fact]
    public void RsaAcceptsSubjectPublicKeyInfoAndPkcs1Keys() {
        using RSA rsa = RSA.Create(2048);
        byte[] data = Encoding.ASCII.GetBytes("signed header data");
        byte[] signature = rsa.SignData(data, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

        OfficePublicKeySignatureCheck spki = Verifier.Verify(OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256,
            rsa.ExportSubjectPublicKeyInfo(), data, signature);
        OfficePublicKeySignatureCheck pkcs1 = Verifier.Verify(OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256,
            rsa.ExportRSAPublicKey(), data, signature);
        OfficePublicKeySignatureCheck wrongHash = Verifier.Verify(OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha1,
            rsa.ExportSubjectPublicKeyInfo(), data, signature);

        Assert.True(spki.IsValid);
        Assert.Equal(2048, spki.KeySizeBits);
        Assert.True(pkcs1.IsValid);
        Assert.Equal(OfficePublicKeySignatureStatus.Invalid, wrongHash.Status);
    }

    [Fact]
    public void MalformedKeysAndSignaturesAreReportedWithoutThrowing() {
        using RSA rsa = RSA.Create(1024);
        using ECDsa ec = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        byte[] data = { 1, 2, 3 };

        Assert.Equal(OfficePublicKeySignatureStatus.MalformedKey,
            Verifier.Verify(OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256, new byte[] { 0x30, 0x03, 0x02, 0x01 }, data, new byte[128]).Status);
        Assert.Equal(OfficePublicKeySignatureStatus.MalformedKey,
            Verifier.Verify(OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256, ec.ExportSubjectPublicKeyInfo(), data, new byte[64]).Status);
        Assert.Equal(OfficePublicKeySignatureStatus.MalformedSignature,
            Verifier.Verify(OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256, rsa.ExportSubjectPublicKeyInfo(), data, new byte[12]).Status);
        Assert.Equal(OfficePublicKeySignatureStatus.MalformedKey,
            Verifier.Verify(OfficePublicKeySignatureAlgorithm.Ed25519, new byte[31], data, new byte[64]).Status);
        Assert.Equal(OfficePublicKeySignatureStatus.MalformedSignature,
            Verifier.Verify(OfficePublicKeySignatureAlgorithm.Ed25519, new byte[32], data, new byte[63]).Status);
        Assert.Equal(OfficePublicKeySignatureStatus.UnsupportedAlgorithm,
            Verifier.Verify((OfficePublicKeySignatureAlgorithm)99, new byte[32], data, new byte[64]).Status);
    }
}
