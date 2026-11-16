using System.Security.Cryptography;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class CoreManagedAesTests {
    private const string Fips197Plaintext = "00112233445566778899AABBCCDDEEFF";

    // FIPS 197 appendix C: one block with each key size. CBC with a zero IV over one block is the bare cipher.
    [Theory]
    [InlineData("000102030405060708090A0B0C0D0E0F", "69C4E0D86A7B0430D8CDB78070B4C55A")]
    [InlineData("000102030405060708090A0B0C0D0E0F1011121314151617", "DDA97CA4864CDFE06EAF70A0EC0D7191")]
    [InlineData("000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F", "8EA2B7CA516745BFEAFC49904B496089")]
    public void BlockCipherMatchesFips197Vectors(string keyHex, string ciphertextHex) {
        byte[] key = Hex(keyHex);
        byte[] iv = new byte[16];

        byte[] encrypted = OfficeManagedAesCryptographyProvider.Default.EncryptCbc(key, iv, Hex(Fips197Plaintext), OfficeAesPadding.None);
        byte[] decrypted = OfficeManagedAesCryptographyProvider.Default.DecryptCbc(key, iv, Hex(ciphertextHex), OfficeAesPadding.None);

        Assert.Equal(Hex(ciphertextHex), encrypted);
        Assert.Equal(Hex(Fips197Plaintext), decrypted);
    }

    // NIST SP 800-38A F.2: CBC-AES128, 192 and 256 over four blocks, in both directions.
    [Theory]
    [InlineData("2B7E151628AED2A6ABF7158809CF4F3C",
        "7649ABAC8119B246CEE98E9B12E9197D5086CB9B507219EE95DB113A917678B273BED6B8E3C1743B7116E69E222295163FF1CAA1681FAC09120ECA307586E1A7")]
    [InlineData("8E73B0F7DA0E6452C810F32B809079E562F8EAD2522C6B7B",
        "4F021DB243BC633D7178183A9FA071E8B4D9ADA9AD7DEDF4E5E738763F69145A571B242012FB7AE07FA9BAAC3DF102E008B0E27988598881D920A9E64F5615CD")]
    [InlineData("603DEB1015CA71BE2B73AEF0857D77811F352C073B6108D72D9810A30914DFF4",
        "F58C4C04D6E5F1BA779EABFB5F7BFBD69CFC4E967EDB808D679F777BC6702C7D39F23369A9D9BACFA530E26304231461B2EB05E2C39BE9FCDA6C19078C6A9D1B")]
    public void CbcMatchesSp80038aVectors(string keyHex, string ciphertextHex) {
        byte[] key = Hex(keyHex);
        byte[] iv = Hex("000102030405060708090A0B0C0D0E0F");
        byte[] plaintext = Hex("6BC1BEE22E409F96E93D7E117393172AAE2D8A571E03AC9C9EB76FAC45AF8E5130C81C46A35CE411E5FBC1191A0A52EFF69F2445DF4F9B17AD2B417BE66C3710");

        Assert.Equal(Hex(ciphertextHex), OfficeManagedAesCryptographyProvider.Default.EncryptCbc(key, iv, plaintext, OfficeAesPadding.None));
        Assert.Equal(plaintext, OfficeManagedAesCryptographyProvider.Default.DecryptCbc(key, iv, Hex(ciphertextHex), OfficeAesPadding.None));
    }

    [Theory]
    [InlineData(16)]
    [InlineData(24)]
    [InlineData(32)]
    public void CbcMatchesThePlatformAesOnRandomInputs(int keyLength) {
        var random = new Random(0x5EED + keyLength);
        using Aes platform = Aes.Create();
        platform.Mode = CipherMode.CBC;
        platform.Padding = PaddingMode.PKCS7;
        for (int iteration = 0; iteration < 64; iteration++) {
            byte[] key = Bytes(random, keyLength);
            byte[] iv = Bytes(random, 16);
            byte[] plaintext = Bytes(random, random.Next(0, 200));
            platform.Key = key;
            platform.IV = iv;
            byte[] expected;
            using (ICryptoTransform encryptor = platform.CreateEncryptor()) {
                expected = encryptor.TransformFinalBlock(plaintext, 0, plaintext.Length);
            }

            byte[] encrypted = OfficeManagedAesCryptographyProvider.Default.EncryptCbc(key, iv, plaintext, OfficeAesPadding.Pkcs7);
            byte[] decrypted = OfficeManagedAesCryptographyProvider.Default.DecryptCbc(key, iv, expected, OfficeAesPadding.Pkcs7);

            Assert.Equal(expected, encrypted);
            Assert.Equal(plaintext, decrypted);
        }
    }

    private static byte[] Bytes(Random random, int length) {
        var bytes = new byte[length];
        random.NextBytes(bytes);
        return bytes;
    }

    private static byte[] Hex(string value) {
        var bytes = new byte[value.Length / 2];
        for (int index = 0; index < bytes.Length; index++) {
            bytes[index] = Convert.ToByte(value.Substring(index * 2, 2), 16);
        }

        return bytes;
    }
}
