using Org.BouncyCastle.Bcpg;
using Org.BouncyCastle.Bcpg.OpenPgp;
using Org.BouncyCastle.Bcpg.Sig;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Security;
using OfficeIMO.Security;

namespace OfficeIMO.Security.Tests;

public sealed class OpenPgpProviderTests {
    private static readonly IOfficeOpenPgpProvider Provider = OfficeSecurityProvider.Default;
    private static readonly byte[] Data = Encoding.ASCII.GetBytes("Signed body\r\n");
    private static readonly Lazy<TestKey> Standard = new(() => TestKey.Create(2048));

    [Fact]
    public void DetachedSignaturesFromPrimaryAndSigningSubkeysVerify() {
        TestKey key = Standard.Value;

        OfficeOpenPgpVerificationResult primary = Provider.VerifyDetached(key.Sign(Data, key.Primary), Data, key.PublicKeys());
        OfficeOpenPgpVerificationResult subkey = Provider.VerifyDetached(key.Sign(Data, key.SigningSubkey), Data, key.PublicKeys());
        OfficeOpenPgpVerificationResult changed = Provider.VerifyDetached(key.Sign(Data, key.Primary), Encoding.ASCII.GetBytes("Signed body!\r\n"), key.PublicKeys());

        Assert.True(primary.IsValid);
        Assert.Equal(primary.Signatures[0].PrimaryKeyFingerprint, primary.Signatures[0].SigningKeyFingerprint);
        Assert.True(subkey.IsValid);
        Assert.NotEqual(subkey.Signatures[0].PrimaryKeyFingerprint, subkey.Signatures[0].SigningKeyFingerprint);
        Assert.Contains("Test Signer <signer@example.com>", subkey.Signatures[0].UserIds);
        Assert.Equal(OfficeOpenPgpSignatureStatus.Invalid, changed.Signatures[0].Status);
    }

    [Fact]
    public void SubkeysNeedAValidBindingAndTheSigningFlag() {
        TestKey key = Standard.Value;

        OfficeOpenPgpVerificationResult encryptionOnly = Provider.VerifyDetached(key.Sign(Data, key.EncryptionSubkey), Data, key.PublicKeys());
        OfficeOpenPgpVerificationResult unbound = Provider.VerifyDetached(key.Sign(Data, key.SigningSubkey), Data,
            new OfficeOpenPgpKeySet(new[] { key.WithoutSigningSubkeyBinding() }));

        Assert.Equal(OfficeOpenPgpSignatureStatus.KeyNotValidForSigning, encryptionOnly.Signatures[0].Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.KeyNotValidForSigning, unbound.Signatures[0].Status);
    }

    [Fact]
    public void RevokedExpiredAndSmallKeysAreRefused() {
        TestKey key = Standard.Value;
        TestKey expiring = TestKey.Create(2048, created: DateTime.UtcNow.AddDays(-3), expiresAfter: TimeSpan.FromDays(1));
        TestKey small = TestKey.Create(1024);

        OfficeOpenPgpVerificationResult revoked = Provider.VerifyDetached(key.Sign(Data, key.Primary), Data,
            new OfficeOpenPgpKeySet(new[] { key.Revoked() }));
        OfficeOpenPgpVerificationResult expired = Provider.VerifyDetached(expiring.Sign(Data, expiring.Primary), Data, expiring.PublicKeys());
        OfficeOpenPgpVerificationResult weak = Provider.VerifyDetached(small.Sign(Data, small.Primary), Data, small.PublicKeys());
        OfficeOpenPgpVerificationResult signatureExpired = Provider.VerifyDetached(
            key.Sign(Data, key.Primary, signedAt: DateTime.UtcNow.AddMinutes(-5), lifetimeSeconds: 60), Data, key.PublicKeys());

        Assert.Equal(OfficeOpenPgpSignatureStatus.KeyRevoked, revoked.Signatures[0].Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.KeyExpired, expired.Signatures[0].Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.WeakAlgorithm, weak.Signatures[0].Status);
        Assert.True(Provider.VerifyDetached(small.Sign(Data, small.Primary), Data, small.PublicKeys(),
            new OfficeOpenPgpOptions { MinimumKeyBits = 1024 }).IsValid);
        Assert.Equal(OfficeOpenPgpSignatureStatus.SignatureExpired, signatureExpired.Signatures[0].Status);
    }

    [Fact]
    public void MalformedInputAndLimitsAreReported() {
        TestKey key = Standard.Value;
        byte[] signature = key.Sign(Data, key.Primary);

        Assert.Equal(OfficeOpenPgpSignatureStatus.Malformed,
            Provider.VerifyDetached(new byte[] { 1, 2, 3 }, Data, key.PublicKeys()).Signatures[0].Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.MissingPublicKey,
            Provider.VerifyDetached(signature, Data, new OfficeOpenPgpKeySet(new[] { new byte[] { 0x99, 0x00 } })).Signatures[0].Status);
        Assert.Empty(Provider.VerifyDetached(signature, Data, key.PublicKeys(), new OfficeOpenPgpOptions { MaxInputBytes = 8 }).Signatures);
        Assert.Equal(OfficeOpenPgpSignatureStatus.Malformed,
            Provider.VerifyCleartext(Encoding.ASCII.GetBytes("-----BEGIN PGP SIGNED MESSAGE-----\n\n-unescaped\n"), key.PublicKeys()).Signatures[0].Status);
    }

    private sealed class TestKey {
        private readonly byte[] _publicRing;

        private TestKey(byte[] publicRing, PgpKeyPair primary, PgpKeyPair signing, PgpKeyPair encryption) {
            _publicRing = publicRing;
            Primary = primary;
            SigningSubkey = signing;
            EncryptionSubkey = encryption;
        }

        internal PgpKeyPair Primary { get; }

        internal PgpKeyPair SigningSubkey { get; }

        internal PgpKeyPair EncryptionSubkey { get; }

        internal static TestKey Create(int bits, DateTime? created = null, TimeSpan? expiresAfter = null) {
            var random = new SecureRandom();
            DateTime time = created ?? DateTime.UtcNow.AddMinutes(-10);
            PgpKeyPair primary = Rsa(bits, time, random);
            PgpKeyPair signing = Rsa(bits, time, random);
            PgpKeyPair encryption = Rsa(bits, time, random);

            var certification = new PgpSignatureSubpacketGenerator();
            certification.SetKeyFlags(false, KeyFlags.CertifyOther | KeyFlags.SignData);
            if (expiresAfter is TimeSpan lifetime) certification.SetKeyExpirationTime(false, (long)lifetime.TotalSeconds);
            var generator = new PgpKeyRingGenerator(PgpSignature.PositiveCertification, primary, "Test Signer <signer@example.com>",
                SymmetricKeyAlgorithmTag.Aes256, Array.Empty<char>(), true, certification.Generate(), null, random);
            var signingFlags = new PgpSignatureSubpacketGenerator();
            signingFlags.SetKeyFlags(false, KeyFlags.SignData);
            signingFlags.SetEmbeddedSignature(false, CrossCertify(signing, primary));
            generator.AddSubKey(signing, signingFlags.Generate(), null);
            var encryptionFlags = new PgpSignatureSubpacketGenerator();
            encryptionFlags.SetKeyFlags(false, KeyFlags.EncryptComms | KeyFlags.EncryptStorage);
            generator.AddSubKey(encryption, encryptionFlags.Generate(), null);
            return new TestKey(generator.GeneratePublicKeyRing().GetEncoded(), primary, signing, encryption);
        }

        internal OfficeOpenPgpKeySet PublicKeys() => new(new[] { PublicRing().GetEncoded() });

        internal byte[] Sign(byte[] data, PgpKeyPair signer, DateTime? signedAt = null, long lifetimeSeconds = 0) {
            var generator = new PgpSignatureGenerator(signer.PublicKey.Algorithm, HashAlgorithmTag.Sha256);
            generator.InitSign(PgpSignature.BinaryDocument, signer.PrivateKey);
            var subpackets = new PgpSignatureSubpacketGenerator();
            subpackets.SetSignatureCreationTime(false, signedAt ?? DateTime.UtcNow);
            subpackets.SetIssuerKeyID(false, signer.KeyId);
            if (lifetimeSeconds > 0) subpackets.SetSignatureExpirationTime(false, lifetimeSeconds);
            generator.SetHashedSubpackets(subpackets.Generate());
            generator.Update(data);
            return generator.Generate().GetEncoded();
        }

        internal byte[] WithoutSigningSubkeyBinding() {
            PgpPublicKeyRing ring = PublicRing();
            PgpPublicKey subkey = ring.GetPublicKey(SigningSubkey.KeyId);
            foreach (PgpSignature binding in subkey.GetSignaturesOfType(PgpSignature.SubkeyBinding).ToList()) {
                subkey = PgpPublicKey.RemoveCertification(subkey, binding);
            }

            return PgpPublicKeyRing.InsertPublicKey(ring, subkey).GetEncoded();
        }

        internal byte[] Revoked() {
            PgpPublicKeyRing ring = PublicRing();
            var generator = new PgpSignatureGenerator(Primary.PublicKey.Algorithm, HashAlgorithmTag.Sha256);
            generator.InitSign(PgpSignature.KeyRevocation, Primary.PrivateKey);
            PgpSignature revocation = generator.GenerateCertification(ring.GetPublicKey());
            return PgpPublicKeyRing.InsertPublicKey(ring, PgpPublicKey.AddCertification(ring.GetPublicKey(), revocation)).GetEncoded();
        }

        private PgpPublicKeyRing PublicRing() => new(_publicRing);

        private static PgpSignature CrossCertify(PgpKeyPair subkey, PgpKeyPair primary) {
            var generator = new PgpSignatureGenerator(subkey.PublicKey.Algorithm, HashAlgorithmTag.Sha256);
            generator.InitSign(PgpSignature.PrimaryKeyBinding, subkey.PrivateKey);
            return generator.GenerateCertification(primary.PublicKey, subkey.PublicKey);
        }

        private static PgpKeyPair Rsa(int bits, DateTime time, SecureRandom random) {
            var generator = new RsaKeyPairGenerator();
            generator.Init(new RsaKeyGenerationParameters(BigInteger.ValueOf(65537), random, bits, 80));
            AsymmetricCipherKeyPair pair = generator.GenerateKeyPair();
            return new PgpKeyPair(PublicKeyAlgorithmTag.RsaGeneral, pair, time);
        }
    }
}
