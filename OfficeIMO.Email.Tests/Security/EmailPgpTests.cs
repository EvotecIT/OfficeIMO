using System.Text.Json;
using Org.BouncyCastle.Bcpg;
using Org.BouncyCastle.Bcpg.OpenPgp;
using Org.BouncyCastle.Security;
using OfficeIMO.Security;

namespace OfficeIMO.Email.Tests;

public sealed class EmailPgpTests {
    private static readonly IOfficeOpenPgpProvider Provider = OfficeSecurityProvider.Default;
    private static readonly Lazy<JsonDocument> Fixtures = new(() => {
        var assembly = typeof(EmailPgpTests).Assembly;
        string name = assembly.GetManifestResourceNames().Single(item => item.EndsWith(".Pgp.pgpy-messages.json", StringComparison.Ordinal));
        using Stream stream = assembly.GetManifestResourceStream(name)!;
        return JsonDocument.Parse(stream);
    });

    private static byte[] Key(string name) => Convert.FromBase64String(Fixtures.Value.RootElement.GetProperty("keys").GetProperty(name).GetString()!);

    private static byte[] MessageBytes(string name) =>
        Convert.FromBase64String(Fixtures.Value.RootElement.GetProperty("messages").GetProperty(name).GetString()!);

    private static EmailDocument Message(string name) => EmailDocument.Load(MessageBytes(name));

    private static OfficeOpenPgpKeySet PublicKeys(params string[] names) => new(names.Select(Key));

    private static OfficeOpenPgpKeySet BobSecret(params string[] publicKeys) => new(publicKeys.Select(Key), new[] { Key("bobSecret") });

    [Theory]
    [InlineData("signed_rsa_sha256", "alicePublic", "Sha256", "Alice Example <alice@example.com>")]
    [InlineData("signed_ed25519_sha512", "bobPublic", "Sha512", "Bob Example <bob@example.net>")]
    public void PgpMimeSignaturesVerifyOverTheExactSignedPart(string name, string key, string hash, string userId) {
        EmailDocument document = Message(name);

        EmailPgpVerificationResult result = EmailPgp.Verify(document, Provider, PublicKeys(key));

        Assert.Equal(EmailProtectionKind.PgpMimeClearSigned, document.Protection.Kind);
        Assert.True(result.IsValid, string.Join("; ", result.Cryptography?.Signatures.Select(s => s.Status + ": " + s.Message) ?? new[] { result.Cryptography?.Message ?? "none" }) + string.Join("|", result.Diagnostics.Select(d => d.Code + d.Message)));
        OfficeOpenPgpSignatureInfo signature = Assert.Single(result.Cryptography!.Signatures);
        Assert.Equal(hash, signature.HashAlgorithm);
        Assert.Equal(4, signature.KeyVersion);
        Assert.Contains(userId, signature.UserIds);
        Assert.StartsWith("Quarterly numbers are attached.", result.SignedContent!.Body.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PGP_MICALG_MISMATCH");
    }

    [Fact]
    public void PgpMimeSignatureFailuresAreTyped() {
        EmailPgpVerificationResult tampered = EmailPgp.Verify(Message("signed_ed25519_tampered"), Provider, PublicKeys("bobPublic"));
        EmailPgpVerificationResult unknownKey = EmailPgp.Verify(Message("signed_ed25519_sha512"), Provider, PublicKeys("alicePublic"));
        EmailPgpVerificationResult weak = EmailPgp.Verify(Message("signed_sha1"), Provider, PublicKeys("bobPublic"));
        EmailPgpVerificationResult weakAllowed = EmailPgp.Verify(Message("signed_sha1"), Provider, PublicKeys("bobPublic"),
            new OfficeOpenPgpOptions { AllowWeakHashAlgorithms = true });
        EmailPgpVerificationResult micalg = EmailPgp.Verify(Message("signed_micalg_mismatch"), Provider, PublicKeys("bobPublic"));

        Assert.Equal(OfficeOpenPgpSignatureStatus.Invalid, tampered.Cryptography!.Signatures[0].Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.MissingPublicKey, unknownKey.Cryptography!.Signatures[0].Status);
        Assert.Equal(OfficeOpenPgpSignatureStatus.WeakAlgorithm, weak.Cryptography!.Signatures[0].Status);
        Assert.True(weakAllowed.IsValid);
        Assert.True(micalg.IsValid);
        Assert.Contains(micalg.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PGP_MICALG_MISMATCH");
        Assert.False(tampered.IsValid || unknownKey.IsValid || weak.IsValid);
    }

    [Fact]
    public void PgpMimeEncryptedMessagesDecryptWithSuppliedSecretKey() {
        EmailDocument document = Message("encrypted_to_bob");

        EmailPgpDecryptionResult result = EmailPgp.Decrypt(document, Provider, BobSecret());

        Assert.Equal(EmailProtectionKind.PgpMimeEncrypted, document.Protection.Kind);
        Assert.True(result.IsDecrypted);
        Assert.True(result.Cryptography!.IsIntegrityProtected);
        Assert.Empty(result.Cryptography.Signatures);
        Assert.False(result.IsSignatureValid);
        Assert.StartsWith("Quarterly numbers are attached.", result.DecryptedContent!.Body.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void SignaturesInsideOrAroundEncryptionAreVerified() {
        EmailPgpDecryptionResult combined = EmailPgp.Decrypt(Message("encrypted_signed_to_bob"), Provider, BobSecret("alicePublic"));
        EmailPgpDecryptionResult combinedUnknownSigner = EmailPgp.Decrypt(Message("encrypted_signed_to_bob"), Provider, BobSecret());
        EmailPgpDecryptionResult nested = EmailPgp.Decrypt(Message("encrypted_nested_signed_to_bob"), Provider, BobSecret("alicePublic"));

        Assert.True(combined.IsSignatureValid);
        Assert.Contains("Alice Example <alice@example.com>", combined.Cryptography!.Signatures[0].UserIds);
        Assert.True(combinedUnknownSigner.IsDecrypted);
        Assert.False(combinedUnknownSigner.IsSignatureValid);
        Assert.Equal(OfficeOpenPgpSignatureStatus.MissingPublicKey, combinedUnknownSigner.Cryptography!.Signatures[0].Status);
        Assert.True(nested.IsDecrypted);
        Assert.True(nested.NestedSignature!.IsValid);
        Assert.True(nested.IsSignatureValid);
    }

    [Fact]
    public void ProtectedSecretKeysNeedTheRightPassphrase() {
        string passphrase = Fixtures.Value.RootElement.GetProperty("keys").GetProperty("alicePassphrase").GetString()!;
        EmailDocument document = Message("encrypted_to_alice");
        byte[][] secret = { Key("aliceSecret") };

        EmailPgpDecryptionResult missing = EmailPgp.Decrypt(document, Provider, new OfficeOpenPgpKeySet(secretKeys: secret));
        EmailPgpDecryptionResult wrong = EmailPgp.Decrypt(document, Provider,
            new OfficeOpenPgpKeySet(secretKeys: secret, passphraseProvider: _ => "wrong".ToCharArray()));
        EmailPgpDecryptionResult right = EmailPgp.Decrypt(document, Provider,
            new OfficeOpenPgpKeySet(secretKeys: secret, passphraseProvider: _ => passphrase.ToCharArray()));
        EmailPgpDecryptionResult otherKey = EmailPgp.Decrypt(document, Provider, BobSecret());

        Assert.Equal(OfficeOpenPgpMessageStatus.PassphraseRequired, missing.Cryptography!.Status);
        Assert.Equal(OfficeOpenPgpMessageStatus.WrongPassphrase, wrong.Cryptography!.Status);
        Assert.True(right.IsDecrypted);
        Assert.Equal(OfficeOpenPgpMessageStatus.NoMatchingSecretKey, otherKey.Cryptography!.Status);
        Assert.Single(otherKey.Cryptography.RecipientKeyIds);
        Assert.Null(missing.DecryptedContent);
    }

    [Fact]
    public void AlteredCiphertextNeverReleasesPlaintext() {
        byte[] binary = Dearmor(PayloadOf(MessageBytes("encrypted_to_bob")));
        binary[binary.Length - 5] ^= 0x01;

        OfficeOpenPgpMessageResult result = Provider.OpenMessage(binary, BobSecret());

        Assert.Equal(OfficeOpenPgpMessageStatus.IntegrityCheckFailed, result.Status);
        Assert.Null(result.Content);
    }

    [Fact]
    public void EncryptionWithoutIntegrityProtectionIsRefusedByDefault() {
        byte[] message = EncryptWithoutIntegrity(Key("bobPublic"), Encoding.ASCII.GetBytes("legacy data"));

        OfficeOpenPgpMessageResult refused = Provider.OpenMessage(message, BobSecret());
        OfficeOpenPgpMessageResult allowed = Provider.OpenMessage(message, BobSecret(),
            new OfficeOpenPgpOptions { AllowMissingIntegrityProtection = true });

        Assert.Equal(OfficeOpenPgpMessageStatus.MissingIntegrityProtection, refused.Status);
        Assert.Null(refused.Content);
        Assert.Equal(OfficeOpenPgpMessageStatus.Opened, allowed.Status);
        Assert.False(allowed.IsIntegrityProtected);
        Assert.Equal("legacy data", Encoding.ASCII.GetString(allowed.Content!));
    }

    [Fact]
    public void InlineCleartextSignaturesAreDetectedAndVerified() {
        EmailDocument document = Message("inline_cleartext_bob");

        EmailPgpVerificationResult result = EmailPgp.Verify(document, Provider, PublicKeys("bobPublic"));
        EmailPgpVerificationResult tampered = EmailPgp.Verify(Message("inline_cleartext_tampered"), Provider, PublicKeys("bobPublic"));
        // RFC 4880 section 7.1: trailing whitespace is not signed, so mail transports may add or strip it.
        byte[] padded = Encoding.ASCII.GetBytes(Encoding.ASCII.GetString(MessageBytes("inline_cleartext_bob")).Replace("Thanks,\r\n", "Thanks, \t \r\n"));
        EmailPgpVerificationResult whitespace = EmailPgp.Verify(EmailDocument.Load(padded), Provider, PublicKeys("bobPublic"));

        Assert.Equal(EmailInlinePgpKind.CleartextSigned, document.Protection.InlinePgp);
        Assert.False(document.Protection.IsProtected);
        Assert.True(result.IsValid);
        Assert.True(result.Cryptography!.Signatures[0].IsTextSignature);
        // Dash-escaping is removed and trailing whitespace, which is not signed, is dropped.
        Assert.Equal("Meeting moved to 10:00.\r\n- dash line\r\nFrom here on.\r\nThanks,\r\nBob", result.SignedText);
        Assert.Equal(OfficeOpenPgpSignatureStatus.Invalid, tampered.Cryptography!.Signatures[0].Status);
        Assert.True(whitespace.IsValid);
        Assert.Equal(result.SignedText, whitespace.SignedText);
    }

    [Fact]
    public void InlineArmoredMessagesAreDecryptedAndTheirSignaturesChecked() {
        EmailDocument document = Message("inline_encrypted_bob");

        EmailPgpDecryptionResult result = EmailPgp.Decrypt(document, Provider, BobSecret("bobPublic"));

        Assert.Equal(EmailInlinePgpKind.Message, document.Protection.InlinePgp);
        Assert.True(result.IsDecrypted);
        Assert.Equal("Inline secret for Bob.", result.DecryptedText);
        Assert.True(result.IsSignatureValid);
    }

    [Fact]
    public void Rfc9580Version4Ed25519LegacyVectorVerifies() {
        // RFC 9580 Appendix A.1 key and A.2 signature over "OpenPGP".
        byte[] key = Encoding.ASCII.GetBytes("-----BEGIN PGP PUBLIC KEY BLOCK-----\n\nxjMEU/NfCxYJKwYBBAHaRw8BAQdAPwmJlL3ZFu1AUxl5NOSofIBzOhKA1i+AEJku\nQ+47JAY=\n-----END PGP PUBLIC KEY BLOCK-----\n");
        byte[] signature = Encoding.ASCII.GetBytes("-----BEGIN PGP SIGNATURE-----\n\niF4EABYIAAYFAlX5X5UACgkQjP3hIZeWWpr2IgD/VvkMypjiECY3vZg/2xbBMd/S\nftgr9N3lYG4NdWrtM2YBANCcT6EVJ/A44PV/IgHYLy6iyQMyZfps60iehUuuYbQE\n-----END PGP SIGNATURE-----\n");

        OfficeOpenPgpVerificationResult valid = Provider.VerifyDetached(signature, Encoding.ASCII.GetBytes("OpenPGP"), new OfficeOpenPgpKeySet(new[] { key }));
        OfficeOpenPgpVerificationResult changed = Provider.VerifyDetached(signature, Encoding.ASCII.GetBytes("OpenPGp"), new OfficeOpenPgpKeySet(new[] { key }));

        Assert.True(valid.IsValid);
        Assert.Equal("C959BDBAFA32A2F89A153B678CFDE12197965A9A", valid.Signatures[0].PrimaryKeyFingerprint);
        Assert.Equal(OfficeOpenPgpSignatureStatus.Invalid, changed.Signatures[0].Status);
    }

    [Fact]
    public void Rfc9580Version6MessagesAreNotReportedAsValid() {
        // RFC 9580 Appendix A.3 certificate and A.6 cleartext signed message (version 6 keys and signatures).
        byte[] certificate = Encoding.ASCII.GetBytes(
            "-----BEGIN PGP PUBLIC KEY BLOCK-----\n\n" +
            "xioGY4d/4xsAAAAg+U2nu0jWCmHlZ3BqZYfQMxmZu52JGggkLq2EVD34laPCsQYf\n" +
            "GwoAAABCBYJjh3/jAwsJBwUVCg4IDAIWAAKbAwIeCSIhBssYbE8GCaaX5NUt+mxy\n" +
            "KwwfHifBilZwj2Ul7Ce62azJBScJAgcCAAAAAK0oIBA+LX0ifsDm185Ecds2v8lw\n" +
            "gyU2kCcUmKfvBXbAf6rhRYWzuQOwEn7E/aLwIwRaLsdry0+VcallHhSu4RN6HWaE\n" +
            "QsiPlR4zxP/TP7mhfVEe7XWPxtnMUMtf15OyA51YBM4qBmOHf+MZAAAAIIaTJINn\n" +
            "+eUBXbki+PSAld2nhJh/LVmFsS+60WyvXkQ1wpsGGBsKAAAALAWCY4d/4wKbDCIh\n" +
            "BssYbE8GCaaX5NUt+mxyKwwfHifBilZwj2Ul7Ce62azJAAAAAAQBIKbpGG2dWTX8\n" +
            "j+VjFM21J0hqWlEg+bdiojWnKfA5AQpWUWtnNwDEM0g12vYxoWM8Y81W+bHBw805\n" +
            "I8kWVkXU6vFOi+HWvv/ira7ofJu16NnoUkhclkUrk0mXubZvyl4GBg==\n" +
            "-----END PGP PUBLIC KEY BLOCK-----\n");
        byte[] cleartext = Encoding.ASCII.GetBytes(
            "-----BEGIN PGP SIGNED MESSAGE-----\n\n" +
            "What we need from the grocery store:\n\n- - tofu\n- - vegetables\n- - noodles\n\n" +
            "-----BEGIN PGP SIGNATURE-----\n\n" +
            "wpgGARsKAAAAKQWCY5ijYyIhBssYbE8GCaaX5NUt+mxyKwwfHifBilZwj2Ul7Ce6\n" +
            "2azJAAAAAGk2IHZJX1AhiJD39eLuPBgiUU9wUA9VHYblySHkBONKU/usJ9BvuAqo\n" +
            "/FvLFuGWMbKAdA+epq7V4HOtAPlBWmU8QOd6aud+aSunHQaaEJ+iTFjP2OMW0KBr\n" +
            "NK2ay45cX1IVAQ==\n" +
            "-----END PGP SIGNATURE-----\n");

        byte[] encrypted = Encoding.ASCII.GetBytes(
            "-----BEGIN PGP MESSAGE-----\n\n" +
            "wV0GIQYSyD8ecG9jCP4VGkF3Q6HwM3kOk+mXhIjR2zeNqZMIhRmHzxjV8bU/gXzO\n" +
            "WgBM85PMiVi93AZfJfhK9QmxfdNnZBjeo1VDeVZheQHgaVf7yopqR6W1FT6NOrfS\n" +
            "aQIHAgZhZBZTW+CwcW1g4FKlbExAf56zaw76/prQoN+bAzxpohup69LA7JW/Vp0l\n" +
            "yZnuSj3hcFj0DfqLTGgr4/u717J+sPWbtQBfgMfG9AOIwwrUBqsFE9zW+f1zdlYo\n" +
            "bhF30A+IitsxxA==\n" +
            "-----END PGP MESSAGE-----\n");

        OfficeOpenPgpVerificationResult result = Provider.VerifyCleartext(cleartext, new OfficeOpenPgpKeySet(new[] { certificate }));
        OfficeOpenPgpMessageResult opened = Provider.OpenMessage(encrypted, new OfficeOpenPgpKeySet());

        Assert.False(result.IsValid);
        Assert.Equal(OfficeOpenPgpSignatureStatus.Unsupported, Assert.Single(result.Signatures).Status);
        Assert.Contains("version 6", result.Signatures[0].Message, StringComparison.Ordinal);
        // RFC 9580 Appendix A.8.5: version 6 PKESK followed by version 2 SEIPD.
        Assert.Equal(OfficeOpenPgpMessageStatus.Unsupported, opened.Status);
        Assert.Null(opened.Content);
    }

    [Fact]
    public void HostileInputIsReportedWithoutThrowing() {
        var random = new Random(3156);
        byte[] source = MessageBytes("signed_ed25519_sha512");
        OfficeOpenPgpKeySet keys = BobSecret("bobPublic");
        for (int iteration = 0; iteration < 200; iteration++) {
            byte[] mutated = (byte[])source.Clone();
            for (int change = 0; change < 1 + iteration % 6; change++) mutated[random.Next(mutated.Length)] = (byte)random.Next(256);
            EmailDocument? document = TryLoad(mutated);
            if (document != null) {
                _ = EmailPgp.Verify(document, Provider, keys);
                _ = EmailPgp.Decrypt(document, Provider, keys);
            }

            _ = Provider.OpenMessage(mutated, keys);
            _ = Provider.VerifyDetached(mutated, mutated, keys);
            _ = Provider.VerifyCleartext(mutated, keys);
        }

        Assert.Equal(OfficeOpenPgpMessageStatus.LimitExceeded,
            Provider.OpenMessage(source, keys, new OfficeOpenPgpOptions { MaxInputBytes = 16 }).Status);
        Assert.Equal(OfficeOpenPgpMessageStatus.Malformed, Provider.OpenMessage(new byte[] { 0x99, 0x01 }, keys).Status);
    }

    // The reader rejects some broken MIME structures outright; that is the reader's contract, not the PGP layer's.
    private static EmailDocument? TryLoad(byte[] data) {
        try {
            return EmailDocument.Load(data);
        } catch (InvalidDataException) {
            return null;
        }
    }

    private static byte[] PayloadOf(byte[] mime) {
        string text = Encoding.ASCII.GetString(mime);
        int start = text.IndexOf("-----BEGIN PGP MESSAGE-----", StringComparison.Ordinal);
        int end = text.IndexOf("-----END PGP MESSAGE-----", StringComparison.Ordinal) + "-----END PGP MESSAGE-----".Length;
        return Encoding.ASCII.GetBytes(text.Substring(start, end - start));
    }

    private static byte[] Dearmor(byte[] armored) {
        using Stream input = PgpUtilities.GetDecoderStream(new MemoryStream(armored));
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] EncryptWithoutIntegrity(byte[] publicKeyBlock, byte[] content) {
        using Stream keyStream = PgpUtilities.GetDecoderStream(new MemoryStream(publicKeyBlock));
        PgpPublicKey encryptionKey = new PgpPublicKeyRingBundle(keyStream).GetKeyRings()
            .SelectMany(ring => ring.GetPublicKeys()).First(key => key.IsEncryptionKey && !key.IsMasterKey);
        var generator = new PgpEncryptedDataGenerator(SymmetricKeyAlgorithmTag.Aes256, withIntegrityPacket: false, new SecureRandom());
        generator.AddMethod(encryptionKey);
        using var output = new MemoryStream();
        using (Stream encrypted = generator.Open(output, new byte[4096])) {
            var literal = new PgpLiteralDataGenerator();
            using Stream literalStream = literal.Open(encrypted, PgpLiteralData.Binary, "data", content.Length, DateTime.UtcNow);
            literalStream.Write(content, 0, content.Length);
        }

        return output.ToArray();
    }
}
