namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Preserves malformed UTF-16 at the JSON boundary so the text inspector can diagnose it.</summary>
internal static class Utf16Transport {
    internal static string? EncodeIfUnpaired(string? text) {
        if (text == null) return null;
        for (int index = 0; index < text.Length; index++) {
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) { index++; continue; }
            if (!char.IsSurrogate(text[index])) continue;
            var bytes = new byte[checked(text.Length * 2)];
            for (int unit = 0; unit < text.Length; unit++) { bytes[unit * 2] = (byte)text[unit]; bytes[unit * 2 + 1] = (byte)(text[unit] >> 8); }
            return Convert.ToBase64String(bytes);
        }
        return null;
    }

    internal static string Decode(string encoded) {
        byte[] bytes = Convert.FromBase64String(encoded);
        if (bytes.Length % 2 != 0) throw new ArgumentException("The UTF-16 option has an incomplete code unit.");
        var units = new char[bytes.Length / 2];
        for (int index = 0; index < units.Length; index++) units[index] = (char)(bytes[index * 2] | (bytes[index * 2 + 1] << 8));
        return new string(units);
    }
}
