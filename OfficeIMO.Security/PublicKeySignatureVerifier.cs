using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Pkcs;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Crypto.Signers;
using Org.BouncyCastle.Security;

namespace OfficeIMO.Security;

/// <summary>Raw RSA and Ed25519 signature checks backed by BouncyCastle.</summary>
internal static class PublicKeySignatureVerifier {
    internal const int MaxRsaKeyBits = 16384;
    private const int MaxEncodedKeyBytes = 4096;
    private const int Ed25519SignatureBytes = 64;

    internal static bool Supports(OfficePublicKeySignatureAlgorithm algorithm) =>
        algorithm == OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256 ||
        algorithm == OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha1 ||
        algorithm == OfficePublicKeySignatureAlgorithm.Ed25519;

    internal static OfficePublicKeySignatureCheck Verify(
        OfficePublicKeySignatureAlgorithm algorithm,
        byte[] publicKey,
        byte[] data,
        byte[] signature) {
        ArgumentNullException.ThrowIfNull(publicKey);
        ArgumentNullException.ThrowIfNull(data);
        ArgumentNullException.ThrowIfNull(signature);
        return algorithm switch {
            OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256 => VerifyRsa(publicKey, data, signature, "SHA-256withRSA"),
            OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha1 => VerifyRsa(publicKey, data, signature, "SHA-1withRSA"),
            OfficePublicKeySignatureAlgorithm.Ed25519 => VerifyEd25519(publicKey, data, signature),
            _ => new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.UnsupportedAlgorithm,
                message: "The signature algorithm is not supported.")
        };
    }

    private static OfficePublicKeySignatureCheck VerifyRsa(byte[] publicKey, byte[] data, byte[] signature, string mechanism) {
        RsaKeyParameters? key = TryDecodeRsaKey(publicKey);
        if (key == null) {
            return new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.MalformedKey,
                message: "The public key is not a DER SubjectPublicKeyInfo or RSAPublicKey RSA key.");
        }

        int bits = key.Modulus.BitLength;
        if (bits > MaxRsaKeyBits || key.Exponent.SignValue <= 0 || !key.Exponent.TestBit(0)) {
            return new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.MalformedKey, bits,
                "The RSA public key is outside the supported bounds.");
        }

        int modulusBytes = (bits + 7) / 8;
        if (signature.Length != modulusBytes) {
            return new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.MalformedSignature, bits,
                "The RSA signature length does not match the key modulus.");
        }

        try {
            ISigner signer = SignerUtilities.GetSigner(mechanism);
            signer.Init(false, key);
            signer.BlockUpdate(data, 0, data.Length);
            return new OfficePublicKeySignatureCheck(signer.VerifySignature(signature)
                ? OfficePublicKeySignatureStatus.Valid
                : OfficePublicKeySignatureStatus.Invalid, bits);
        } catch (Exception ex) when (ex is CryptoException or ArgumentException or InvalidOperationException) {
            return new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.Invalid, bits,
                "The RSA signature could not be processed.");
        }
    }

    private static RsaKeyParameters? TryDecodeRsaKey(byte[] encoded) {
        if (encoded.Length == 0 || encoded.Length > MaxEncodedKeyBytes) return null;
        Asn1Sequence? sequence;
        try {
            sequence = Asn1Object.FromByteArray(encoded) as Asn1Sequence;
        } catch (Exception ex) when (ex is IOException or ArgumentException or InvalidOperationException or InvalidCastException) {
            return null;
        }

        if (sequence == null) return null;
        try {
            if (sequence.Count == 2 && sequence[0] is DerInteger) {
                RsaPublicKeyStructure pkcs1 = RsaPublicKeyStructure.GetInstance(sequence);
                return new RsaKeyParameters(false, pkcs1.Modulus, pkcs1.PublicExponent);
            }

            SubjectPublicKeyInfo info = SubjectPublicKeyInfo.GetInstance(sequence);
            if (!PkcsObjectIdentifiers.RsaEncryption.Equals(info.Algorithm.Algorithm)) return null;
            RsaPublicKeyStructure rsa = RsaPublicKeyStructure.GetInstance(info.ParsePublicKey());
            return new RsaKeyParameters(false, rsa.Modulus, rsa.PublicExponent);
        } catch (Exception ex) when (ex is IOException or ArgumentException or InvalidOperationException or InvalidCastException) {
            return null;
        }
    }

    private static OfficePublicKeySignatureCheck VerifyEd25519(byte[] publicKey, byte[] data, byte[] signature) {
        if (publicKey.Length != Ed25519PublicKeyParameters.KeySize) {
            return new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.MalformedKey,
                message: "The Ed25519 public key must be 32 bytes.");
        }

        if (signature.Length != Ed25519SignatureBytes) {
            return new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.MalformedSignature, 256,
                "The Ed25519 signature must be 64 bytes.");
        }

        try {
            var key = new Ed25519PublicKeyParameters(publicKey, 0);
            var signer = new Ed25519Signer();
            signer.Init(false, key);
            signer.BlockUpdate(data, 0, data.Length);
            return new OfficePublicKeySignatureCheck(signer.VerifySignature(signature)
                ? OfficePublicKeySignatureStatus.Valid
                : OfficePublicKeySignatureStatus.Invalid, 256);
        } catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) {
            return new OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus.MalformedKey,
                message: "The Ed25519 public key is not a valid curve point.");
        }
    }
}
