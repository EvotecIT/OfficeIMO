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
    private static readonly Lazy<TestKey> WithoutBackSignature = new(() => TestKey.Create(2048, crossCertify: false));

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

    [Fact]
    public void RevocationsAndSubkeysMustBeAuthenticated() {
        TestKey key = Standard.Value;
        TestKey other = WithoutBackSignature.Value;

        OfficeOpenPgpVerificationResult forgedRevocation = Provider.VerifyDetached(key.Sign(Data, key.Primary), Data,
            new OfficeOpenPgpKeySet(new[] { key.Revoked(other.Primary) }));
        OfficeOpenPgpVerificationResult noBackSignature = Provider.VerifyDetached(other.Sign(Data, other.SigningSubkey), Data, other.PublicKeys());
        OfficeOpenPgpVerificationResult staleCopyFirst = Provider.VerifyDetached(key.Sign(Data, key.Primary), Data,
            new OfficeOpenPgpKeySet(new[] { key.PublicKeys().PublicKeys[0], key.Revoked() }));

        Assert.True(forgedRevocation.IsValid);
        Assert.Equal(OfficeOpenPgpSignatureStatus.KeyNotValidForSigning, noBackSignature.Signatures[0].Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.KeyRevoked, staleCopyFirst.Signatures[0].Status);
    }

    [Fact]
    public void HiddenRecipientsTryEveryKeyAndPassphrasesAreUtf8() {
        TestKey key = Standard.Value;
        var keys = new OfficeOpenPgpKeySet(secretKeys: new[] { key.SecretRing }, passphraseProvider: _ => TestKey.Passphrase.ToCharArray());

        OfficeOpenPgpMessageResult named = Provider.OpenMessage(key.Encrypt(Data, hideRecipient: false), keys);
        OfficeOpenPgpMessageResult hidden = Provider.OpenMessage(key.Encrypt(Data, hideRecipient: true), keys);

        Assert.Equal(OfficeOpenPgpMessageStatus.Opened, named.Status);
        Assert.Equal(OfficeOpenPgpMessageStatus.Opened, hidden.Status);
        Assert.Equal(Data, hidden.Content);
    }

    [Fact]
    public void DecompressedPacketsAndTruncatedMarkersStayWithinLimits() {
        TestKey key = Standard.Value;
        byte[] signature = key.Sign(Data, key.Primary);
        using var compressedOutput = new MemoryStream();
        var compressor = new PgpCompressedDataGenerator(CompressionAlgorithmTag.Zip);
        using (Stream compressed = compressor.Open(compressedOutput)) {
            for (int index = 0; index < 200; index++) compressed.Write(signature, 0, signature.Length);
        }

        OfficeOpenPgpMessageResult bomb = Provider.OpenMessage(compressedOutput.ToArray(), key.PublicKeys(),
            new OfficeOpenPgpOptions { MaxSignatures = 1000, MaxOutputBytes = 2048 });
        OfficeOpenPgpVerificationResult marker = Provider.VerifyDetached(new byte[] { 0xCA, 0xFF, 0x7F, 0xFF, 0xFF, 0xFF, 0x50 }, Data, key.PublicKeys());

        Assert.Equal(OfficeOpenPgpMessageStatus.LimitExceeded, bomb.Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.Malformed, marker.Signatures[0].Status);
    }

    private sealed class TestKey {
        private readonly byte[] _publicRing;

        internal const string Passphrase = "p\u00e4ss";

        private TestKey(byte[] publicRing, byte[] secretRing, PgpKeyPair primary, PgpKeyPair signing, PgpKeyPair encryption) {
            _publicRing = publicRing;
            SecretRing = secretRing;
            Primary = primary;
            SigningSubkey = signing;
            EncryptionSubkey = encryption;
        }

        internal PgpKeyPair Primary { get; }

        internal PgpKeyPair SigningSubkey { get; }

        internal PgpKeyPair EncryptionSubkey { get; }

        internal byte[] SecretRing { get; }

        internal static TestKey Create(int bits, DateTime? created = null, TimeSpan? expiresAfter = null, bool crossCertify = true) {
            var random = new SecureRandom();
            DateTime time = created ?? DateTime.UtcNow.AddMinutes(-10);
            PgpKeyPair primary = Rsa(bits, time, random);
            PgpKeyPair signing = Rsa(bits, time, random);
            PgpKeyPair encryption = Rsa(bits, time, random);

            var certification = new PgpSignatureSubpacketGenerator();
            certification.SetKeyFlags(false, KeyFlags.CertifyOther | KeyFlags.SignData);
            if (expiresAfter is TimeSpan lifetime) certification.SetKeyExpirationTime(false, (long)lifetime.TotalSeconds);
            var generator = new PgpKeyRingGenerator(PgpSignature.PositiveCertification, primary, "Test Signer <signer@example.com>",
                SymmetricKeyAlgorithmTag.Aes256, true, Passphrase.ToCharArray(), true, certification.Generate(), null, random);
            var signingFlags = new PgpSignatureSubpacketGenerator();
            signingFlags.SetKeyFlags(false, KeyFlags.SignData);
            if (crossCertify) signingFlags.SetEmbeddedSignature(false, CrossCertify(signing, primary));
            generator.AddSubKey(signing, signingFlags.Generate(), null);
            var encryptionFlags = new PgpSignatureSubpacketGenerator();
            encryptionFlags.SetKeyFlags(false, KeyFlags.EncryptComms | KeyFlags.EncryptStorage);
            generator.AddSubKey(encryption, encryptionFlags.Generate(), null);
            return new TestKey(generator.GeneratePublicKeyRing().GetEncoded(), generator.GenerateSecretKeyRing().GetEncoded(),
                primary, signing, encryption);
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

        internal byte[] Revoked(PgpKeyPair? signer = null) {
            PgpPublicKeyRing ring = PublicRing();
            signer ??= Primary;
            var generator = new PgpSignatureGenerator(signer.PublicKey.Algorithm, HashAlgorithmTag.Sha256);
            generator.InitSign(PgpSignature.KeyRevocation, signer.PrivateKey);
            PgpSignature revocation = generator.GenerateCertification(ring.GetPublicKey());
            return PgpPublicKeyRing.InsertPublicKey(ring, PgpPublicKey.AddCertification(ring.GetPublicKey(), revocation)).GetEncoded();
        }

        internal byte[] Encrypt(byte[] data, bool hideRecipient) {
            var generator = new PgpEncryptedDataGenerator(SymmetricKeyAlgorithmTag.Aes256, true, new SecureRandom());
            generator.AddMethod(EncryptionSubkey.PublicKey);
            using var output = new MemoryStream();
            using (Stream encrypted = generator.Open(output, new byte[4096])) {
                var literal = new PgpLiteralDataGenerator();
                using Stream plain = literal.Open(encrypted, PgpLiteralData.Binary, "message.txt", data.Length, DateTime.UtcNow);
                plain.Write(data, 0, data.Length);
            }

            byte[] message = output.ToArray();
            if (hideRecipient) {
                // RFC 4880 section 5.1: a key ID of zero hides the recipient. It follows the version octet (3).
                byte[] keyId = BitConverter.GetBytes(EncryptionSubkey.KeyId);
                if (BitConverter.IsLittleEndian) Array.Reverse(keyId);
                int position = IndexOf(message, new byte[] { 3 }.Concat(keyId).ToArray());
                Assert.True(position >= 0);
                Array.Clear(message, position + 1, keyId.Length);
            }

            return message;
        }

        private static int IndexOf(byte[] data, byte[] pattern) {
            for (int index = 0; index + pattern.Length <= data.Length; index++) {
                if (data.AsSpan(index, pattern.Length).SequenceEqual(pattern)) return index;
            }

            return -1;
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
