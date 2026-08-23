namespace OfficeIMO.Security;

/// <summary>Padding modes supported by the dependency-free OfficeIMO AES contract.</summary>
public enum OfficeAesPadding {
    /// <summary>The input length must be an exact multiple of the AES block size.</summary>
    None,

    /// <summary>PKCS #7 padding is added during encryption and validated during decryption.</summary>
    Pkcs7
}

/// <summary>
/// Supplies synchronous AES-CBC operations for hosts where the platform cryptography implementation is unavailable.
/// Implementations are expected to return a new byte array and must not mutate caller-owned inputs.
/// </summary>
public interface IOfficeAesCryptographyProvider {
    /// <summary>Gets a stable provider name for diagnostics.</summary>
    string Name { get; }

    /// <summary>Encrypts the supplied plaintext using AES-CBC.</summary>
    byte[] EncryptCbc(byte[] key, byte[] initializationVector, byte[] plaintext, OfficeAesPadding padding);

    /// <summary>Decrypts the supplied ciphertext using AES-CBC.</summary>
    byte[] DecryptCbc(byte[] key, byte[] initializationVector, byte[] ciphertext, OfficeAesPadding padding);
}
