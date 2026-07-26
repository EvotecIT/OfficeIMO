using System.Security.Cryptography;

namespace OfficeIMO.Email;

internal static class EmailHashing {
    internal static string ComputeSha256HexLower(string value) {
        using SHA256 sha256 = SHA256.Create();
        return ToHexLower(sha256.ComputeHash(Encoding.UTF8.GetBytes(value)));
    }

    internal static string ToHexLower(byte[] value) =>
        BitConverter.ToString(value).Replace("-", string.Empty).ToLowerInvariant();
}
