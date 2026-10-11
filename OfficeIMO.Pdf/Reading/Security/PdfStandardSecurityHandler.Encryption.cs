using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfStandardSecurityHandler {
    internal PdfObject EncryptObject(int objectNumber, int generation, PdfObject value, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (value is PdfStringObj text) {
            return new PdfStringObj(EncryptData(objectNumber, generation, text.RawBytes, _stringMethod), text.UseTextStringEncoding) {
                HasIncompleteSyntax = text.HasIncompleteSyntax
            };
        }

        if (value is PdfArray array) {
            var encrypted = new PdfArray { HasIncompleteSyntax = array.HasIncompleteSyntax };
            for (int i = 0; i < array.Items.Count; i++) {
                encrypted.Items.Add(EncryptObject(objectNumber, generation, array.Items[i], cancellationToken));
            }

            return encrypted;
        }

        if (value is PdfDictionary dictionary) {
            return EncryptDictionary(objectNumber, generation, dictionary, cancellationToken);
        }

        if (value is PdfStream stream) {
            bool skipData = ShouldSkipStreamData(stream.Dictionary);
            PdfDictionary encryptedDictionary = EncryptDictionary(objectNumber, generation, stream.Dictionary, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] encryptedData = skipData
                ? stream.Data
                : EncryptData(objectNumber, generation, stream.Data, _streamMethod);
            return new PdfStream(encryptedDictionary, encryptedData, stream.DecodingFailed, stream.DecodingError) {
                HasIncompleteSyntax = stream.HasIncompleteSyntax
            };
        }

        return value;
    }

    private PdfDictionary EncryptDictionary(int objectNumber, int generation, PdfDictionary dictionary, CancellationToken cancellationToken) {
        var encrypted = new PdfDictionary { HasIncompleteSyntax = dictionary.HasIncompleteSyntax };
        foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            encrypted.Items[item.Key] = IsSignatureContents(dictionary, item.Key)
                ? item.Value
                : EncryptObject(objectNumber, generation, item.Value, cancellationToken);
        }

        return encrypted;
    }

    private byte[] EncryptData(int objectNumber, int generation, byte[] data, PdfCryptMethod method) {
        if (method == PdfCryptMethod.Identity) {
            return data;
        }

        if (method == PdfCryptMethod.AesV3) {
            return EncryptAesCbc(_fileKey, data);
        }

        byte[] objectKey = ComputeObjectKey(objectNumber, generation, method == PdfCryptMethod.AesV2);
        if (method == PdfCryptMethod.Rc4) {
            return Rc4.Transform(objectKey, data);
        }

        if (method == PdfCryptMethod.AesV2) {
            return EncryptAesCbc(objectKey, data);
        }

        throw new PdfUnsupportedEncryptionException("Unsupported PDF crypt filter method.");
    }

    private byte[] EncryptAesCbc(byte[] key, byte[] data) {
        var iv = new byte[16];
        using (RandomNumberGenerator random = RandomNumberGenerator.Create()) {
            random.GetBytes(iv);
        }

        byte[] ciphertext = PdfAesCryptography.EncryptPkcs7(key, iv, data, _aesCryptographyProvider);
        var result = new byte[iv.Length + ciphertext.Length];
        Buffer.BlockCopy(iv, 0, result, 0, iv.Length);
        Buffer.BlockCopy(ciphertext, 0, result, iv.Length, ciphertext.Length);
        return result;
    }
}
