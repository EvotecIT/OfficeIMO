using System.Security.Cryptography;

namespace OfficeIMO.Epub;

internal static class EpubFontObfuscation {
    private const int IdpfPrefixLength = 1040;
    private const int AdobePrefixLength = 1024;

    internal static bool TryDeobfuscate(byte[] data, EpubEncryptionKind kind, string? identifier,
        out byte[] result, out string? error) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        result = Array.Empty<byte>();
        error = null;
        byte[]? key;
        int prefixLength;
        switch (kind) {
            case EpubEncryptionKind.IdpfFontObfuscation:
                key = CreateIdpfKey(identifier);
                prefixLength = IdpfPrefixLength;
                if (key == null) error = "the package unique identifier is missing";
                break;
            case EpubEncryptionKind.AdobeFontObfuscation:
                key = CreateAdobeKey(identifier);
                prefixLength = AdobePrefixLength;
                if (key == null) error = "the package unique identifier is not a UUID";
                break;
            default:
                key = null;
                prefixLength = 0;
                error = "the declared algorithm is not a recognized font-obfuscation profile";
                break;
        }
        if (key == null) return false;

        result = (byte[])data.Clone();
        int count = Math.Min(prefixLength, result.Length);
        for (int index = 0; index < count; index++) result[index] ^= key[index % key.Length];
        return true;
    }

    private static byte[]? CreateIdpfKey(string? identifier) {
        if (identifier == null) return null;
        var normalized = new StringBuilder(identifier!.Length);
        foreach (char character in identifier) {
            if (character != '\u0020' && character != '\u0009' &&
                character != '\u000D' && character != '\u000A') {
                normalized.Append(character);
            }
        }
        if (normalized.Length == 0) return null;
        using var sha1 = SHA1.Create();
        return sha1.ComputeHash(Encoding.UTF8.GetBytes(normalized.ToString()));
    }

    private static byte[]? CreateAdobeKey(string? identifier) {
        if (string.IsNullOrWhiteSpace(identifier)) return null;
        string value = identifier!.Trim();
        if (value.StartsWith("urn:uuid:", StringComparison.OrdinalIgnoreCase)) value = value.Substring(9);
        value = value.Trim('{', '}').Replace("-", string.Empty);
        if (value.Length != 32) return null;
        var key = new byte[16];
        for (int index = 0; index < key.Length; index++) {
            if (!TryParseHex(value[index * 2], value[index * 2 + 1], out key[index])) return null;
        }
        return key;
    }

    private static bool TryParseHex(char high, char low, out byte value) {
        int highValue = HexValue(high);
        int lowValue = HexValue(low);
        if (highValue < 0 || lowValue < 0) {
            value = 0;
            return false;
        }
        value = (byte)((highValue << 4) | lowValue);
        return true;
    }

    private static int HexValue(char value) {
        if (value >= '0' && value <= '9') return value - '0';
        if (value >= 'a' && value <= 'f') return value - 'a' + 10;
        if (value >= 'A' && value <= 'F') return value - 'A' + 10;
        return -1;
    }
}
