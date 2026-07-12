namespace OfficeIMO.Email;

/// <summary>Normalizes caller-controlled header fields before artifact serialization.</summary>
internal static class MimeHeaderSafety {
    internal static string SanitizeName(string value) {
        var result = new StringBuilder(value.Length);
        foreach (char character in value) {
            if ((character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z') ||
                (character >= '0' && character <= '9') || character == '-') result.Append(character);
        }
        return result.Length == 0 ? "X-OfficeIMO-Header" : result.ToString();
    }

    internal static string SanitizeValue(string value) {
        return value.Replace("\r", string.Empty).Replace("\n", " ");
    }
}
