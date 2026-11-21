using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

#pragma warning disable CA1850 // Static HashData is unavailable on netstandard2.0 and net472.
internal static partial class PdfStandardSecurityWriter {
    /// <summary>
    /// Public-key security (ISO 32000-2 7.6.5), /SubFilter /adbe.pkcs7.s5 with an AESV3 crypt filter. One CMS
    /// EnvelopedData carries the 20-byte seed and the four permission bytes (most significant first) to every recipient;
    /// the file key is SHA-256 over the seed, the recipient entries and, when metadata stays clear, four 0xFF bytes.
    /// </summary>
    private static PdfEncryptionAssembly EncryptPublicKey(
        IReadOnlyList<byte[]> sourceObjects,
        PdfStandardEncryptionOptions options,
        long objectMemoryLimitBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfPublicKeyEncryptionOptions publicKey = options.PublicKey!;
        byte[] seed = options.PublicKeySeed!;
        byte[] recipientContent = CreatePublicKeyRecipientContent(seed, options.Permissions);
        byte[] envelope;
        try {
            envelope = publicKey.SecurityProvider.EncryptCms(recipientContent, publicKey.RecipientCertificates)
                ?? throw new InvalidOperationException("The security provider returned no CMS EnvelopedData.");
        } finally {
            Array.Clear(recipientContent, 0, recipientContent.Length);
        }

        cancellationToken.ThrowIfCancellationRequested();
        byte[] fileKey = DerivePublicKeyFileKey(seed, envelope, options.EncryptMetadata);
        byte[] fileId = CreateFileId();
        int encryptionObjectNumber = sourceObjects.Count + 1;
        var objects = new PdfObjectStore(objectMemoryLimitBytes);
        try {
            for (int i = 0; i < sourceObjects.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                objects.Add(EncryptAesIndirectObject(
                    sourceObjects[i],
                    i + 1,
                    fileKey,
                    options.EncryptMetadata,
                    deriveObjectKey: false,
                    options.AesCryptographyProvider,
                    cancellationToken));
            }

            cancellationToken.ThrowIfCancellationRequested();
            objects.Add(PdfObjectBytes.WrapIndirectObject(
                encryptionObjectNumber,
                BuildPublicKeyEncryptionDictionary(envelope, options.EncryptMetadata)));
            return new PdfEncryptionAssembly(objects, encryptionObjectNumber, fileId, objects);
        } catch {
            objects.Dispose();
            throw;
        } finally {
            Array.Clear(fileKey, 0, fileKey.Length);
        }
    }

    internal static byte[] CreatePublicKeyRecipientContent(byte[] seed, int permissions) {
        var content = new byte[24];
        Buffer.BlockCopy(seed, 0, content, 0, 20);
        // Bit 1 grants nothing the reader enforces; PDFBox and pyHanko set it in recipient data for Acrobat.
        permissions |= 1;
        unchecked {
            content[20] = (byte)(permissions >> 24);
            content[21] = (byte)(permissions >> 16);
            content[22] = (byte)(permissions >> 8);
            content[23] = (byte)permissions;
        }

        return content;
    }

    private static byte[] DerivePublicKeyFileKey(byte[] seed, byte[] envelope, bool encryptMetadata) {
        using SHA256 sha = SHA256.Create();
        sha.TransformBlock(seed, 0, seed.Length, null, 0);
        sha.TransformBlock(envelope, 0, envelope.Length, null, 0);
        if (!encryptMetadata) {
            byte[] metadataMarker = { 0xFF, 0xFF, 0xFF, 0xFF };
            sha.TransformBlock(metadataMarker, 0, metadataMarker.Length, null, 0);
        }

        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return sha.Hash!;
    }

    private static string BuildPublicKeyEncryptionDictionary(byte[] envelope, bool encryptMetadata) {
        string metadata = encryptMetadata ? "true" : "false";
        return "<< /Filter /Adobe.PubSec /SubFilter /adbe.pkcs7.s5 /V 5 /Length 256" +
            " /EncryptMetadata " + metadata +
            " /CF << /DefaultCryptFilter << /AuthEvent /DocOpen /CFM /AESV3 /Length 256" +
            " /EncryptMetadata " + metadata +
            " /Recipients [" + PdfSyntaxEscaper.HexString(envelope) + "] >> >>" +
            " /StmF /DefaultCryptFilter /StrF /DefaultCryptFilter >>";
    }
}
#pragma warning restore CA1850
