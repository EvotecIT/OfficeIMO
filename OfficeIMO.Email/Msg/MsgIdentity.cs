using System.Security.Cryptography;

namespace OfficeIMO.Email;

/// <summary>Creates deterministic MAPI identifiers and one-off address-book entries.</summary>
internal static class MsgIdentity {
    private static readonly byte[] OneOffProviderUid = {
        0x81, 0x2B, 0x1F, 0xA4, 0xBE, 0xA3, 0x10, 0x19,
        0x9D, 0x6E, 0x00, 0xDD, 0x01, 0x0F, 0x54, 0x02
    };

    internal static byte[] CreateStableBytes(string scope, int length, params string?[] values) {
        if (length <= 0 || length > 32) throw new ArgumentOutOfRangeException(nameof(length));
        string input = string.Concat(scope, "\n", string.Join("\n", values.Select(value => value ?? string.Empty)));
        using SHA256 sha256 = SHA256.Create();
        byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
        byte[] result = new byte[length];
        Buffer.BlockCopy(hash, 0, result, 0, length);
        return result;
    }

    internal static byte[] CreateSearchKey(string? addressType, string? address) {
        string value = string.Concat(
            string.IsNullOrWhiteSpace(addressType) ? "SMTP" : addressType!.Trim().ToUpperInvariant(),
            ":",
            address?.Trim().ToUpperInvariant() ?? string.Empty,
            "\0");
        return Encoding.ASCII.GetBytes(value);
    }

    internal static byte[] CreateOneOffEntryId(EmailAddress address) {
        using var stream = new MemoryStream();
        stream.Write(new byte[4], 0, 4);
        stream.Write(OneOffProviderUid, 0, OneOffProviderUid.Length);
        stream.WriteByte(0);
        stream.WriteByte(0);
        stream.WriteByte(0x09);
        stream.WriteByte(0xE8);
        WriteUnicodeString(stream, address.DisplayName ?? address.Address ?? string.Empty);
        WriteUnicodeString(stream, string.IsNullOrWhiteSpace(address.AddressType) ? "SMTP" : address.AddressType!);
        WriteUnicodeString(stream, address.Address ?? string.Empty);
        return stream.ToArray();
    }

    private static void WriteUnicodeString(Stream stream, string value) {
        byte[] bytes = Encoding.Unicode.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
        stream.WriteByte(0);
        stream.WriteByte(0);
    }
}
