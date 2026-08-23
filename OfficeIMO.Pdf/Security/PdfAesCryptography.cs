using System.Security.Cryptography;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>Central AES-CBC owner for PDF Standard security writing and reading.</summary>
internal static class PdfAesCryptography {
    internal static byte[] EncryptNoPadding(
        byte[] key,
        byte[] initializationVector,
        byte[] plaintext,
        IOfficeAesCryptographyProvider? provider) =>
        Transform(key, initializationVector, plaintext, OfficeAesPadding.None, provider, encrypt: true);

    internal static byte[] DecryptNoPadding(
        byte[] key,
        byte[] initializationVector,
        byte[] ciphertext,
        IOfficeAesCryptographyProvider? provider) =>
        Transform(key, initializationVector, ciphertext, OfficeAesPadding.None, provider, encrypt: false);

    internal static byte[] EncryptPkcs7(
        byte[] key,
        byte[] initializationVector,
        byte[] plaintext,
        IOfficeAesCryptographyProvider? provider) =>
        Transform(key, initializationVector, plaintext, OfficeAesPadding.Pkcs7, provider, encrypt: true);

    private static byte[] Transform(
        byte[] key,
        byte[] initializationVector,
        byte[] input,
        OfficeAesPadding padding,
        IOfficeAesCryptographyProvider? provider,
        bool encrypt) {
        ValidateInputs(key, initializationVector, input, padding, encrypt);
        if (provider != null) {
            byte[] output = encrypt
                ? provider.EncryptCbc(key, initializationVector, input, padding)
                : provider.DecryptCbc(key, initializationVector, input, padding);
            return output ?? throw new InvalidOperationException(
                $"AES provider '{provider.Name}' returned no output.");
        }

        try {
            using Aes aes = Aes.Create();
            aes.Mode = CipherMode.CBC;
            aes.Padding = padding == OfficeAesPadding.Pkcs7 ? PaddingMode.PKCS7 : PaddingMode.None;
            aes.Key = key;
            aes.IV = initializationVector;
            using ICryptoTransform transform = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor();
            return transform.TransformFinalBlock(input, 0, input.Length);
        } catch (PlatformNotSupportedException exception) {
            throw new PlatformNotSupportedException(
                "This host does not provide synchronous AES-CBC. Supply an IOfficeAesCryptographyProvider, such as OfficeManagedAesCryptographyProvider from OfficeIMO.Core.",
                exception);
        }
    }

    private static void ValidateInputs(
        byte[] key,
        byte[] initializationVector,
        byte[] input,
        OfficeAesPadding padding,
        bool encrypt) {
        Guard.NotNull(key, nameof(key));
        Guard.NotNull(initializationVector, nameof(initializationVector));
        Guard.NotNull(input, nameof(input));
        if (key.Length != 16 && key.Length != 24 && key.Length != 32) {
            throw new ArgumentException("AES keys must contain 16, 24, or 32 bytes.", nameof(key));
        }
        if (initializationVector.Length != 16) {
            throw new ArgumentException("AES-CBC initialization vectors must contain 16 bytes.", nameof(initializationVector));
        }
        if (padding == OfficeAesPadding.None && (input.Length % 16) != 0) {
            throw new ArgumentException("Unpadded AES-CBC input must be an exact multiple of 16 bytes.", nameof(input));
        }
        if (!encrypt && padding == OfficeAesPadding.Pkcs7 && (input.Length == 0 || (input.Length % 16) != 0)) {
            throw new ArgumentException("PKCS #7 AES-CBC ciphertext must contain complete AES blocks.", nameof(input));
        }
    }
}
