using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.IWork.Internal;

/// <summary>Derives stable UUIDv5 names from the complete normalized model, separating document, theme, layout and slide identities.</summary>
internal static class IWorkKeynoteIdentity {
    private static readonly byte[] UrlNamespace = { 0x6b, 0xa7, 0xb8, 0x11, 0x9d, 0xad, 0x11, 0xd1,
        0x80, 0xb4, 0x00, 0xc0, 0x4f, 0xd4, 0x30, 0xc8 };

    internal static (string Text, ulong Lower, ulong Upper) Create(byte[] modelHash, string role) {
        var hex = new StringBuilder(64);
        foreach (byte value in modelHash) hex.Append(value.ToString("x2", CultureInfo.InvariantCulture));
        byte[] name = Encoding.UTF8.GetBytes("https://github.com/EvotecIT/OfficeIMO/keynote/" + hex + "/" + role);
        var input = new byte[UrlNamespace.Length + name.Length];
        Buffer.BlockCopy(UrlNamespace, 0, input, 0, UrlNamespace.Length);
        Buffer.BlockCopy(name, 0, input, UrlNamespace.Length, name.Length);
        // SHA-1 is required by UUIDv5; this identifier is not used as an integrity or security digest.
        using var sha = SHA1.Create();
        byte[] hash = sha.ComputeHash(input);
        hash[6] = (byte)((hash[6] & 15) | 0x50);
        hash[8] = (byte)((hash[8] & 63) | 0x80);
        var text = new StringBuilder(36);
        for (int index = 0; index < 16; index++) {
            if (index == 4 || index == 6 || index == 8 || index == 10) text.Append('-');
            text.Append(hash[index].ToString("x2", CultureInfo.InvariantCulture));
        }
        return (text.ToString(), ReadNetworkUInt64(hash, 8), ReadNetworkUInt64(hash, 0));
    }

    private static ulong ReadNetworkUInt64(byte[] bytes, int start) {
        ulong result = 0;
        for (int index = start; index < start + 8; index++) result = result << 8 | bytes[index];
        return result;
    }
}
