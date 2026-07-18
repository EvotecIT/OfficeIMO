using System.Security.Cryptography;

namespace OfficeIMO.Email;

/// <summary>Creates deterministic MAPI identifiers and one-off address-book entries.</summary>
internal static class MsgIdentity {
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

    internal static byte[] CreateOneOffEntryId(EmailAddress address) => OutlookEntryIdCodec.EncodeOneOff(address);
}
