using OfficeIMO.Security;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeManagedAesCryptographyProviderTests {
    [Theory]
    [InlineData(
        "2B7E151628AED2A6ABF7158809CF4F3C",
        "7649ABAC8119B246CEE98E9B12E9197D5086CB9B507219EE95DB113A917678B273BED6B8E3C1743B7116E69E222295163FF1CAA1681FAC09120ECA307586E1A7")]
    [InlineData(
        "8E73B0F7DA0E6452C810F32B809079E562F8EAD2522C6B7B",
        "4F021DB243BC633D7178183A9FA071E8B4D9ADA9AD7DEDF4E5E738763F69145A571B242012FB7AE07FA9BAAC3DF102E008B0E27988598881D920A9E64F5615CD")]
    [InlineData(
        "603DEB1015CA71BE2B73AEF0857D77811F352C073B6108D72D9810A30914DFF4",
        "F58C4C04D6E5F1BA779EABFB5F7BFBD69CFC4E967EDB808D679F777BC6702C7D39F23369A9D9BACFA530E26304231461B2EB05E2C39BE9FCDA6C19078C6A9D1B")]
    public void AesCbcNoPadding_MatchesNistSp80038aVectors(string keyHex, string expectedHex) {
        byte[] key = Convert.FromHexString(keyHex);
        byte[] iv = Convert.FromHexString("000102030405060708090A0B0C0D0E0F");
        byte[] plaintext = Convert.FromHexString(
            "6BC1BEE22E409F96E93D7E117393172A" +
            "AE2D8A571E03AC9C9EB76FAC45AF8E51" +
            "30C81C46A35CE411E5FBC1191A0A52EF" +
            "F69F2445DF4F9B17AD2B417BE66C3710");
        byte[] expected = Convert.FromHexString(expectedHex);

        byte[] encrypted = OfficeManagedAesCryptographyProvider.Default.EncryptCbc(
            key,
            iv,
            plaintext,
            OfficeAesPadding.None);
        byte[] decrypted = OfficeManagedAesCryptographyProvider.Default.DecryptCbc(
            key,
            iv,
            encrypted,
            OfficeAesPadding.None);

        Assert.Equal(expected, encrypted);
        Assert.Equal(plaintext, decrypted);
    }

    [Fact]
    public void AesCbcPkcs7_RoundTripsNonBlockAlignedPayloadWithoutMutatingInputs() {
        byte[] key = Convert.FromHexString("603DEB1015CA71BE2B73AEF0857D7781" +
                                           "1F352C073B6108D72D9810A30914DFF4");
        byte[] iv = Convert.FromHexString("000102030405060708090A0B0C0D0E0F");
        byte[] plaintext = System.Text.Encoding.UTF8.GetBytes("OfficeIMO browser AES-256");
        byte[] keySnapshot = (byte[])key.Clone();
        byte[] ivSnapshot = (byte[])iv.Clone();
        byte[] plaintextSnapshot = (byte[])plaintext.Clone();

        byte[] encrypted = OfficeManagedAesCryptographyProvider.Default.EncryptCbc(
            key,
            iv,
            plaintext,
            OfficeAesPadding.Pkcs7);
        byte[] decrypted = OfficeManagedAesCryptographyProvider.Default.DecryptCbc(
            key,
            iv,
            encrypted,
            OfficeAesPadding.Pkcs7);

        Assert.Equal(0, encrypted.Length % 16);
        Assert.Equal(plaintext, decrypted);
        Assert.Equal(keySnapshot, key);
        Assert.Equal(ivSnapshot, iv);
        Assert.Equal(plaintextSnapshot, plaintext);
    }

    [Fact]
    public void AesCbcPkcs7_RejectsInvalidPadding() {
        byte[] key = new byte[32];
        byte[] iv = new byte[16];
        byte[] invalidPaddedPlaintext = new byte[16];
        byte[] ciphertext = OfficeManagedAesCryptographyProvider.Default.EncryptCbc(
            key,
            iv,
            invalidPaddedPlaintext,
            OfficeAesPadding.None);

        Assert.Throws<CryptographicException>(() =>
            OfficeManagedAesCryptographyProvider.Default.DecryptCbc(
                key,
                iv,
                ciphertext,
                OfficeAesPadding.Pkcs7));
    }

    [Fact]
    public void AesCbcNoPadding_AcceptsEmptyPayload() {
        byte[] encrypted = OfficeManagedAesCryptographyProvider.Default.EncryptCbc(
            new byte[16],
            new byte[16],
            Array.Empty<byte>(),
            OfficeAesPadding.None);

        Assert.Empty(encrypted);
        Assert.Empty(OfficeManagedAesCryptographyProvider.Default.DecryptCbc(
            new byte[16],
            new byte[16],
            encrypted,
            OfficeAesPadding.None));
    }
}
