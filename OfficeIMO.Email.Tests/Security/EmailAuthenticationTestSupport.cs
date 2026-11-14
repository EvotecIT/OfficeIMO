using System.Security.Cryptography;
using System.Text.Json;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Crypto.Signers;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using OfficeIMO.Security;

namespace OfficeIMO.Email.Tests;

internal static class EmailAuthenticationTestSupport {
    internal static readonly IOfficePublicKeySignatureVerifier Verifier = OfficeSecurityProvider.Default;

    internal static byte[] ReadFixture(string name) {
        var assembly = typeof(EmailAuthenticationTestSupport).Assembly;
        string resource = assembly.GetManifestResourceNames().Single(item => item.EndsWith(".Authentication." + name, StringComparison.Ordinal));
        using Stream stream = assembly.GetManifestResourceStream(resource)!;
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    internal static JsonDocument ReadJsonFixture(string name) => JsonDocument.Parse(ReadFixture(name));

    internal static byte[] Ascii(string text) => Encoding.ASCII.GetBytes(text.Replace("\r\n", "\n").Replace("\n", "\r\n"));
}

/// <summary>Test-only DKIM signer over BouncyCastle keys, used to exercise policy paths the vectors do not cover.</summary>
internal sealed class TestDkimSigner {
    private readonly AsymmetricKeyParameter _privateKey;
    private readonly bool _ed25519;

    private TestDkimSigner(AsymmetricKeyParameter privateKey, string record, bool ed25519) {
        _privateKey = privateKey;
        KeyRecord = record;
        _ed25519 = ed25519;
    }

    internal string KeyRecord { get; }

    internal static TestDkimSigner Rsa(int bits, string recordSuffix = "") {
        var generator = new RsaKeyPairGenerator();
        generator.Init(new RsaKeyGenerationParameters(Org.BouncyCastle.Math.BigInteger.ValueOf(65537), new SecureRandom(), bits, 80));
        AsymmetricCipherKeyPair pair = generator.GenerateKeyPair();
        byte[] spki = SubjectPublicKeyInfoFactory.CreateSubjectPublicKeyInfo(pair.Public).GetDerEncoded();
        return new TestDkimSigner(pair.Private, "v=DKIM1; k=rsa; p=" + Convert.ToBase64String(spki) + recordSuffix, false);
    }

    internal static TestDkimSigner Ed25519(string recordSuffix = "") {
        var generator = new Ed25519KeyPairGenerator();
        generator.Init(new Ed25519KeyGenerationParameters(new SecureRandom()));
        AsymmetricCipherKeyPair pair = generator.GenerateKeyPair();
        byte[] key = ((Ed25519PublicKeyParameters)pair.Public).GetEncoded();
        return new TestDkimSigner(pair.Private, "v=DKIM1; k=ed25519; p=" + Convert.ToBase64String(key) + recordSuffix, true);
    }

    /// <summary>Prepends a DKIM-Signature built from <paramref name="tags"/> (without bh= and b=) to the message.</summary>
    internal byte[] Sign(byte[] message, string tags, EmailDkimCanonicalization header = EmailDkimCanonicalization.Relaxed,
        EmailDkimCanonicalization body = EmailDkimCanonicalization.Relaxed, long? length = null, bool computeBodyHash = true) {
        EmailAuthenticationMessage parsed = EmailAuthenticationMessage.Parse(message, 1000, out _);
        byte[] canonicalBody = EmailDkimCanonicalizer.CanonicalizeBody(parsed.Body, body);
        using SHA256 sha256 = SHA256.Create();
        byte[] bodyHash = sha256.ComputeHash(canonicalBody, 0, (int)(length ?? canonicalBody.Length));
        if (!computeBodyHash) bodyHash = new byte[32];
        string headerValue = " " + tags + "; c=" + Name(header) + "/" + Name(body) + "; bh=" + Convert.ToBase64String(bodyHash) + "; b=";
        string unsignedField = "DKIM-Signature:" + headerValue + "\r\n";

        var data = new StringBuilder();
        string signedNames = tags.Split(';').Select(item => item.Trim()).First(item => item.StartsWith("h=", StringComparison.Ordinal)).Substring(2);
        var used = new HashSet<int>();
        foreach (string name in signedNames.Split(':').Select(item => item.Trim())) {
            for (int index = parsed.Fields.Count - 1; index >= 0; index--) {
                if (used.Contains(index) || !parsed.Fields[index].Is(name)) continue;
                used.Add(index);
                data.Append(EmailDkimCanonicalizer.CanonicalizeHeader(parsed.Fields[index].Raw, header));
                break;
            }
        }

        string canonicalSignature = EmailDkimCanonicalizer.CanonicalizeHeader(unsignedField, header);
        data.Append(canonicalSignature.Substring(0, canonicalSignature.Length - 2));
        byte[] signedBytes = EmailAuthenticationMessage.ToBytes(data.ToString());
        byte[] signature;
        if (_ed25519) {
            var signer = new Ed25519Signer();
            signer.Init(true, _privateKey);
            byte[] digest = sha256.ComputeHash(signedBytes);
            signer.BlockUpdate(digest, 0, digest.Length);
            signature = signer.GenerateSignature();
        } else {
            ISigner signer = SignerUtilities.GetSigner("SHA-256withRSA");
            signer.Init(true, _privateKey);
            signer.BlockUpdate(signedBytes, 0, signedBytes.Length);
            signature = signer.GenerateSignature();
        }

        byte[] field = Encoding.ASCII.GetBytes("DKIM-Signature:" + headerValue + Convert.ToBase64String(signature) + "\r\n");
        return field.Concat(message).ToArray();
    }

    private static string Name(EmailDkimCanonicalization value) => value == EmailDkimCanonicalization.Relaxed ? "relaxed" : "simple";
}
