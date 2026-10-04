#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using BcKeyTransRecipientInfo = Org.BouncyCastle.Asn1.Cms.KeyTransRecipientInfo;
using BcRecipientInfo = Org.BouncyCastle.Asn1.Cms.RecipientInfo;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.X509;
using AlgorithmIdentifier = Org.BouncyCastle.Asn1.X509.AlgorithmIdentifier;
using Org.BouncyCastle.Cms;

namespace OfficeIMO.Security;

/// <summary>Uses the supplied platform RSA handle without exporting its private parameters or searching stores.</summary>
internal static class PlatformCmsEnvelopeDecryptor {
    internal static bool Supports(CmsEnvelopedData envelope, RecipientInformation recipient) {
        if (recipient is not KeyTransRecipientInformation ||
            envelope.EncryptionAlgOid is not ("2.16.840.1.101.3.4.1.2" or "2.16.840.1.101.3.4.1.22" or "2.16.840.1.101.3.4.1.42")) return false;
        AlgorithmIdentifier algorithm = recipient.KeyEncryptionAlgorithmID;
        if (algorithm.Algorithm.Id == "1.2.840.113549.1.1.1") return IsNullParameter(algorithm.Parameters);
        if (algorithm.Algorithm.Id != "1.2.840.113549.1.1.7" || algorithm.Parameters is not Org.BouncyCastle.Asn1.Asn1Sequence) return false;
        var oaep = Org.BouncyCastle.Asn1.Pkcs.RsaesOaepParameters.GetInstance(algorithm.Parameters);
        AlgorithmIdentifier hash = oaep.HashAlgorithm;
        if (hash.Algorithm.Id is not ("1.3.14.3.2.26" or "2.16.840.1.101.3.4.2.1" or "2.16.840.1.101.3.4.2.2" or "2.16.840.1.101.3.4.2.3") ||
            !IsNullParameter(hash.Parameters) || oaep.MaskGenAlgorithm.Algorithm.Id != "1.2.840.113549.1.1.8" ||
            oaep.MaskGenAlgorithm.Parameters is not Org.BouncyCastle.Asn1.Asn1Sequence ||
            oaep.PSourceAlgorithm.Algorithm.Id != "1.2.840.113549.1.1.9") return false;
        AlgorithmIdentifier mgfHash = AlgorithmIdentifier.GetInstance(oaep.MaskGenAlgorithm.Parameters);
        return mgfHash.Algorithm.Equals(hash.Algorithm) && IsNullParameter(mgfHash.Parameters) &&
            (oaep.PSourceAlgorithm.Parameters == null ||
             oaep.PSourceAlgorithm.Parameters is Org.BouncyCastle.Asn1.Asn1OctetString label && label.GetOctets().Length == 0);
    }

    private static bool IsNullParameter(Org.BouncyCastle.Asn1.Asn1Encodable? parameter) =>
        parameter == null || parameter.ToAsn1Object() is Org.BouncyCastle.Asn1.DerNull;

    internal static byte[] Decrypt(CmsEnvelopedData envelope, RecipientInformation recipient, RSA privateKey, long maximumBytes) {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumBytes);
        if (!Supports(envelope, recipient)) throw new NotSupportedException("The envelope is outside the platform RSA/AES-CBC contract.");
        // PKCS#7 padding consumes between one and sixteen bytes. Reject oversized content before the platform
        // parser can allocate plaintext; the final check covers the remaining one-block ambiguity.
        long ciphertextBytes = envelope.EnvelopedData.EncryptedContentInfo.EncryptedContent.GetOctets().LongLength;
        if (ciphertextBytes - 16 > maximumBytes)
            throw new SecurityContentLimitExceededException(ciphertextBytes - 16, maximumBytes);
        BcKeyTransRecipientInfo? encodedRecipient = null;
        foreach (Org.BouncyCastle.Asn1.Asn1Encodable candidate in envelope.EnvelopedData.RecipientInfos) {
            var info = BcRecipientInfo.GetInstance(candidate);
            if (info.Info is not BcKeyTransRecipientInfo keyTransport) continue;
            var id = new RecipientID();
            if (keyTransport.RecipientIdentifier.IsTagged) {
                id.SubjectKeyIdentifier = SubjectKeyIdentifier.GetInstance(keyTransport.RecipientIdentifier.ID).GetEncoded();
            } else {
                var issuer = IssuerAndSerialNumber.GetInstance(keyTransport.RecipientIdentifier.ID);
                id.Issuer = issuer.Issuer;
                id.SerialNumber = issuer.SerialNumber.Value;
            }
            if (id.Equals(recipient.RecipientID)) { encodedRecipient = keyTransport; break; }
        }
        if (encodedRecipient == null) throw new CryptographicException("The selected CMS RSA recipient is absent.");
        byte[] encryptedKey = encodedRecipient.EncryptedKey.GetOctets();
        byte[]? parameters = encodedRecipient.KeyEncryptionAlgorithm.Parameters?.GetEncoded();
        var platform = new EnvelopedCms();
        platform.Decode(envelope.GetEncoded());
        foreach (System.Security.Cryptography.Pkcs.RecipientInfo candidate in platform.RecipientInfos) {
            if (candidate is not System.Security.Cryptography.Pkcs.KeyTransRecipientInfo keyTransport ||
                keyTransport.KeyEncryptionAlgorithm.Oid.Value != encodedRecipient.KeyEncryptionAlgorithm.Algorithm.Id ||
                !keyTransport.EncryptedKey.SequenceEqual(encryptedKey) ||
                !EquivalentParameters(parameters, keyTransport.KeyEncryptionAlgorithm.Parameters)) continue;
            platform.Decrypt(candidate, privateKey);
            byte[] content = platform.ContentInfo.Content;
            if (content.LongLength > maximumBytes) {
                CryptographicOperations.ZeroMemory(content);
                throw new SecurityContentLimitExceededException(content.LongLength, maximumBytes);
            }
            return content;
        }
        throw new CryptographicException("The platform parser could not resolve the selected CMS recipient.");
    }

    private static bool EquivalentParameters(byte[]? encoded, byte[] platform) =>
        IsAbsentOrNull(encoded) && IsAbsentOrNull(platform) || encoded != null && encoded.SequenceEqual(platform);

    private static bool IsAbsentOrNull(byte[]? value) => value == null || value.Length == 0 ||
        value.Length == 2 && value[0] == 5 && value[1] == 0;
}
#endif
