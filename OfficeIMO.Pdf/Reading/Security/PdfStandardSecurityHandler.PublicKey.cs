using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

// Public-key security handler (ISO 32000-1, 7.6.4; ISO 32000-2, 7.6.5). The CMS envelope is opened by the
// caller's IOfficeSecurityProvider; this file only selects recipients and derives the PDF file key.
#pragma warning disable CA1850 // Static HashData is unavailable on netstandard2.0 and net472.
internal sealed partial class PdfStandardSecurityHandler {
    internal const string PublicKeyFilterName = "Adobe.PubSec";
    private const int PublicKeySeedLength = 20;
    private const int PublicKeyRecipientContentLength = 24;
    private const int MaxPublicKeyRecipients = 256;
    private const long MaxPublicKeyRecipientBytes = 1024L * 1024L;
    private const string RecipientNotFoundFindingCode = "EnvelopeRecipientNotFound";

    private static PdfStandardSecurityHandler CreatePublicKey(
        PdfDictionary encryptionDictionary,
        PdfPublicKeyDecryptionOptions? publicKeyDecryption,
        IOfficeAesCryptographyProvider? aesCryptographyProvider,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string subFilter = encryptionDictionary.Get<PdfName>("SubFilter")?.Name ?? string.Empty;
        int version = GetRequiredInt(encryptionDictionary, "V");
        PublicKeyLayout layout = subFilter switch {
            "adbe.pkcs7.s3" or "adbe.pkcs7.s4" => ResolveDocumentLevelPublicKeyLayout(encryptionDictionary, subFilter, version),
            "adbe.pkcs7.s5" => ResolveCryptFilterPublicKeyLayout(encryptionDictionary, version),
            _ => throw new PdfUnsupportedEncryptionException(
                "Unsupported public-key security handler /SubFilter /" + (subFilter.Length == 0 ? "(missing)" : subFilter) +
                ". Supported values are adbe.pkcs7.s3, adbe.pkcs7.s4 and adbe.pkcs7.s5.")
        };

        if (publicKeyDecryption is null) {
            throw new PdfRecipientCertificateRequiredException(
                "This PDF is encrypted for certificate recipients. Set PdfLoadOptions.PublicKeyDecryption with a security provider and a recipient certificate that has its private key.");
        }

        byte[] recipientContent = RecoverPublicKeyRecipientContent(layout.Recipients, publicKeyDecryption, cancellationToken);
        try {
            int permissions = (recipientContent[20] << 24) | (recipientContent[21] << 16) |
                (recipientContent[22] << 8) | recipientContent[23];
            byte[] fileKey = DerivePublicKeyFileKey(recipientContent, layout, cancellationToken);
            return new PdfStandardSecurityHandler(
                fileKey,
                revision: 0,
                layout.KeyLengthBytes,
                layout.StreamMethod,
                layout.StringMethod,
                layout.EncryptMetadata,
                PdfPasswordAuthenticationRole.PublicKeyRecipient,
                aesCryptographyProvider,
                permissions);
        } finally {
            Array.Clear(recipientContent, 0, recipientContent.Length);
        }
    }

    private static PublicKeyLayout ResolveDocumentLevelPublicKeyLayout(PdfDictionary encryptionDictionary, string subFilter, int version) {
        if (version != 1 && version != 2) {
            throw new PdfUnsupportedEncryptionException(
                "Public-key /SubFilter /" + subFilter + " requires /V 1 or 2; this file declares /V " + version + ".");
        }

        if (encryptionDictionary.Get<PdfBoolean>("EncryptMetadata") is { Value: false }) {
            throw new PdfUnsupportedEncryptionException(
                "/EncryptMetadata false is only defined for crypt-filter encryption (/V 4 or 5).");
        }

        int keyLengthBits = version == 1 ? 40 : GetOptionalInt(encryptionDictionary, "Length") ?? 40;
        int keyLengthBytes = ValidateRc4KeyLength(keyLengthBits);
        IReadOnlyList<byte[]> recipients = ReadPublicKeyRecipients(encryptionDictionary, "the encryption dictionary");
        return new PublicKeyLayout(recipients, PdfCryptMethod.Rc4, PdfCryptMethod.Rc4, keyLengthBytes, encryptMetadata: true, useSha256: false);
    }

    private static PublicKeyLayout ResolveCryptFilterPublicKeyLayout(PdfDictionary encryptionDictionary, int version) {
        if (version != 4 && version != 5) {
            throw new PdfUnsupportedEncryptionException(
                "Public-key /SubFilter /adbe.pkcs7.s5 requires /V 4 or 5; this file declares /V " + version + ".");
        }

        string streamFilterName = encryptionDictionary.Get<PdfName>("StmF")?.Name ?? "Identity";
        string stringFilterName = encryptionDictionary.Get<PdfName>("StrF")?.Name ?? "Identity";
        bool streamIdentity = string.Equals(streamFilterName, "Identity", StringComparison.Ordinal);
        bool stringIdentity = string.Equals(stringFilterName, "Identity", StringComparison.Ordinal);
        if (streamIdentity && stringIdentity) {
            throw new PdfUnsupportedEncryptionException("The public-key encryption dictionary does not select an encrypting crypt filter.");
        }

        if (!streamIdentity && !stringIdentity && !string.Equals(streamFilterName, stringFilterName, StringComparison.Ordinal)) {
            // Each public-key crypt filter carries its own recipients and therefore its own file key.
            throw new PdfUnsupportedEncryptionException(
                "Public-key PDFs that use different crypt filters for streams (/" + streamFilterName + ") and strings (/" +
                stringFilterName + ") are not supported.");
        }

        string filterName = streamIdentity ? stringFilterName : streamFilterName;
        PdfDictionary filter = encryptionDictionary.Get<PdfDictionary>("CF")?.Get<PdfDictionary>(filterName)
            ?? throw new PdfUnsupportedEncryptionException("The public-key crypt filter /" + filterName + " is not defined in /CF.");
        string cfm = filter.Get<PdfName>("CFM")?.Name ?? "None";
        PdfCryptMethod method;
        int keyLengthBytes;
        switch (cfm) {
            case "V2" when version == 4:
                method = PdfCryptMethod.Rc4;
                keyLengthBytes = ValidateRc4KeyLength(ResolveCryptFilterKeyLengthBits(encryptionDictionary, filter) ?? 128);
                break;
            case "AESV2" when version == 4:
                method = PdfCryptMethod.AesV2;
                keyLengthBytes = 16;
                break;
            case "AESV3" when version == 5:
                method = PdfCryptMethod.AesV3;
                keyLengthBytes = 32;
                break;
            default:
                throw new PdfUnsupportedEncryptionException(
                    "Unsupported public-key crypt filter method /" + cfm + " for /V " + version +
                    ". Supported combinations are /V 4 with /V2 or /AESV2 and /V 5 with /AESV3.");
        }

        bool? filterMetadata = filter.Get<PdfBoolean>("EncryptMetadata")?.Value;
        bool? documentMetadata = encryptionDictionary.Get<PdfBoolean>("EncryptMetadata")?.Value;
        if (filterMetadata.HasValue && documentMetadata.HasValue && filterMetadata.Value != documentMetadata.Value) {
            throw new PdfUnsupportedEncryptionException(
                "The public-key crypt filter and the encryption dictionary disagree on /EncryptMetadata.");
        }

        bool encryptMetadata = filterMetadata ?? documentMetadata ?? true;
        IReadOnlyList<byte[]> recipients = ReadPublicKeyRecipients(filter, "crypt filter /" + filterName);
        return new PublicKeyLayout(
            recipients,
            streamIdentity ? PdfCryptMethod.Identity : method,
            stringIdentity ? PdfCryptMethod.Identity : method,
            keyLengthBytes,
            encryptMetadata,
            useSha256: method == PdfCryptMethod.AesV3);
    }

    private static int? ResolveCryptFilterKeyLengthBits(PdfDictionary encryptionDictionary, PdfDictionary filter) {
        int? documentBits = GetOptionalInt(encryptionDictionary, "Length");
        if (documentBits.HasValue) {
            return documentBits;
        }

        // Writers disagree on whether a crypt filter /Length is in bytes or bits; values up to 16 can only be bytes.
        int? filterLength = GetOptionalInt(filter, "Length");
        return filterLength is > 0 and <= 16 ? filterLength * 8 : filterLength;
    }

    private static int ValidateRc4KeyLength(int keyLengthBits) {
        if (keyLengthBits < 40 || keyLengthBits > 128 || keyLengthBits % 8 != 0) {
            throw new PdfUnsupportedEncryptionException(
                "Public-key RC4 key length must be a multiple of 8 between 40 and 128 bits; this file declares " + keyLengthBits + ".");
        }

        return keyLengthBits / 8;
    }

    private static List<byte[]> ReadPublicKeyRecipients(PdfDictionary dictionary, string location) {
        PdfObject? value = dictionary.Items.TryGetValue("Recipients", out PdfObject? raw) ? raw : null;
        var recipients = new List<byte[]>();
        if (value is PdfStringObj single) {
            recipients.Add(single.RawBytes);
        } else if (value is PdfArray array) {
            if (array.Items.Count > MaxPublicKeyRecipients) {
                throw new PdfUnsupportedEncryptionException(
                    "The public-key /Recipients array in " + location + " exceeds " + MaxPublicKeyRecipients + " entries.");
            }

            foreach (PdfObject item in array.Items) {
                if (item is not PdfStringObj text) {
                    throw new PdfUnsupportedEncryptionException("The public-key /Recipients array in " + location + " contains a non-string entry.");
                }

                recipients.Add(text.RawBytes);
            }
        }

        if (recipients.Count == 0) {
            throw new PdfUnsupportedEncryptionException("The public-key /Recipients entry in " + location + " is missing or empty.");
        }

        foreach (byte[] recipient in recipients) {
            if (recipient.Length == 0 || recipient.LongLength > MaxPublicKeyRecipientBytes) {
                throw new PdfUnsupportedEncryptionException(
                    "A public-key recipient entry in " + location + " is empty or larger than " + MaxPublicKeyRecipientBytes + " bytes.");
            }
        }

        return recipients;
    }

    private static byte[] RecoverPublicKeyRecipientContent(
        IReadOnlyList<byte[]> recipients,
        PdfPublicKeyDecryptionOptions options,
        CancellationToken cancellationToken) {
        var envelopeOptions = new CmsEnvelopeOptions {
            MaxContentBytes = 4096,
            MaxEncodedBytes = MaxPublicKeyRecipientBytes,
            MaxRecipients = MaxPublicKeyRecipients
        };
        var failures = new List<string>();
        foreach (byte[] recipient in recipients) {
            foreach (X509Certificate2 certificate in options.RecipientCertificates) {
                cancellationToken.ThrowIfCancellationRequested();
                CmsDecryptionResult result;
                try {
                    result = options.SecurityProvider.DecryptCms(recipient, certificate, envelopeOptions);
                } catch (Exception exception) when (exception is not OperationCanceledException &&
                    exception is not OutOfMemoryException && exception is not StackOverflowException) {
                    failures.Add(exception.GetType().Name + ": " + exception.Message);
                    continue;
                }

                if (result.Decrypted && result.Content is not null) {
                    if (result.Content.Length != PublicKeyRecipientContentLength) {
                        Array.Clear(result.Content, 0, result.Content.Length);
                        throw new PdfUnsupportedEncryptionException(
                            "The public-key recipient data must be a 20-byte seed followed by 4 permission bytes.");
                    }

                    return result.Content;
                }

                bool notThisRecipient = result.Parsed && result.Findings.Any(static finding =>
                    string.Equals(finding.Code, RecipientNotFoundFindingCode, StringComparison.Ordinal));
                if (!notThisRecipient) {
                    failures.Add(string.Join("; ", result.Findings.Select(static finding => finding.Code + ": " + finding.Message)));
                }
            }
        }

        if (failures.Count > 0) {
            throw new PdfRecipientDecryptionException(
                "A supplied certificate could not open the PDF public-key recipient data: " + failures[0]);
        }

        throw new PdfRecipientNotFoundException(
            "None of the supplied certificates matches a recipient of this public-key encrypted PDF.");
    }

    private static byte[] DerivePublicKeyFileKey(byte[] recipientContent, PublicKeyLayout layout, CancellationToken cancellationToken) {
        using HashAlgorithm hash = layout.UseSha256 ? SHA256.Create() : CreateSha1();
        hash.TransformBlock(recipientContent, 0, PublicKeySeedLength, null, 0);
        foreach (byte[] recipient in layout.Recipients) {
            cancellationToken.ThrowIfCancellationRequested();
            hash.TransformBlock(recipient, 0, recipient.Length, null, 0);
        }

        if (!layout.EncryptMetadata) {
            byte[] metadataMarker = { 0xFF, 0xFF, 0xFF, 0xFF };
            hash.TransformBlock(metadataMarker, 0, metadataMarker.Length, null, 0);
        }

        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return Take(hash.Hash!, layout.KeyLengthBytes, cancellationToken);
    }

    private static SHA1 CreateSha1() {
#pragma warning disable CA5350 // ISO 32000 public-key key derivation for /V 1, 2 and 4 is defined with SHA-1.
        return SHA1.Create();
#pragma warning restore CA5350
    }

    private sealed class PublicKeyLayout {
        internal PublicKeyLayout(
            IReadOnlyList<byte[]> recipients,
            PdfCryptMethod streamMethod,
            PdfCryptMethod stringMethod,
            int keyLengthBytes,
            bool encryptMetadata,
            bool useSha256) {
            Recipients = recipients;
            StreamMethod = streamMethod;
            StringMethod = stringMethod;
            KeyLengthBytes = keyLengthBytes;
            EncryptMetadata = encryptMetadata;
            UseSha256 = useSha256;
        }

        internal IReadOnlyList<byte[]> Recipients { get; }
        internal PdfCryptMethod StreamMethod { get; }
        internal PdfCryptMethod StringMethod { get; }
        internal int KeyLengthBytes { get; }
        internal bool EncryptMetadata { get; }
        internal bool UseSha256 { get; }
    }
}
#pragma warning restore CA1850
