using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfStandardSecurityHandler {
    internal PdfObject EncryptObject(int objectNumber, int generation, PdfObject value) {
        if (value is PdfStringObj text) {
            return new PdfStringObj(EncryptData(objectNumber, generation, text.RawBytes, _stringMethod), text.UseTextStringEncoding);
        }

        if (value is PdfArray array) {
            var encrypted = new PdfArray();
            for (int i = 0; i < array.Items.Count; i++) {
                encrypted.Items.Add(EncryptObject(objectNumber, generation, array.Items[i]));
            }

            return encrypted;
        }

        if (value is PdfDictionary dictionary) {
            return EncryptDictionary(objectNumber, generation, dictionary);
        }

        if (value is PdfStream stream) {
            bool skipData = ShouldSkipStreamData(stream.Dictionary);
            PdfDictionary encryptedDictionary = EncryptDictionary(objectNumber, generation, stream.Dictionary);
            byte[] encryptedData = skipData
                ? stream.Data
                : EncryptData(objectNumber, generation, stream.Data, _streamMethod);
            return new PdfStream(encryptedDictionary, encryptedData, stream.DecodingFailed, stream.DecodingError);
        }

        return value;
    }

    private PdfDictionary EncryptDictionary(int objectNumber, int generation, PdfDictionary dictionary) {
        var encrypted = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
            encrypted.Items[item.Key] = EncryptObject(objectNumber, generation, item.Value);
        }

        return encrypted;
    }

    private byte[] EncryptData(int objectNumber, int generation, byte[] data, PdfCryptMethod method) {
        if (method == PdfCryptMethod.Identity || data.Length == 0) {
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

    private static byte[] EncryptAesCbc(byte[] key, byte[] data) {
        var iv = new byte[16];
        using (RandomNumberGenerator random = RandomNumberGenerator.Create()) {
            random.GetBytes(iv);
        }

        using Aes aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.Key = key;
        aes.IV = iv;
        using ICryptoTransform encryptor = aes.CreateEncryptor();
        byte[] ciphertext = encryptor.TransformFinalBlock(data, 0, data.Length);
        var result = new byte[iv.Length + ciphertext.Length];
        Buffer.BlockCopy(iv, 0, result, 0, iv.Length);
        Buffer.BlockCopy(ciphertext, 0, result, iv.Length, ciphertext.Length);
        return result;
    }
}
