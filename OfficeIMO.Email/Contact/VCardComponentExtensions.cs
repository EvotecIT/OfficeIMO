namespace OfficeIMO.Email;

/// <summary>Typed text and group helpers over the lossless vCard property model.</summary>
public static class VCardComponentExtensions {
    /// <summary>Returns the first property value with vCard text escaping decoded.</summary>
    public static string? GetVCardText(this ContentLineComponent card, string propertyName) {
        if (card == null) throw new ArgumentNullException(nameof(card));
        ContentLineProperty? property = card.GetFirstProperty(propertyName);
        return property == null ? null : DecodeText(property.Value);
    }

    /// <summary>Creates or replaces one escaped text property.</summary>
    public static ContentLineProperty SetVCardText(this ContentLineComponent card, string propertyName,
        string value) {
        if (card == null) throw new ArgumentNullException(nameof(card));
        return card.SetProperty(propertyName, EncodeText(value ?? throw new ArgumentNullException(nameof(value))));
    }

    /// <summary>Adds a repeated escaped text property.</summary>
    public static ContentLineProperty AddVCardText(this ContentLineComponent card, string propertyName,
        string value) {
        if (card == null) throw new ArgumentNullException(nameof(card));
        return card.AddProperty(propertyName, EncodeText(value ?? throw new ArgumentNullException(nameof(value))));
    }

    /// <summary>Decodes vCard TEXT escaping.</summary>
    public static string DecodeText(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        var result = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (value[index] != '\\' || index + 1 >= value.Length) {
                result.Append(value[index]);
                continue;
            }
            char escaped = value[++index];
            if (escaped == 'n' || escaped == 'N') result.Append('\n');
            else if (escaped == '\\' || escaped == ',' || escaped == ';') result.Append(escaped);
            else result.Append('\\').Append(escaped);
        }
        return result.ToString();
    }

    /// <summary>Encodes a vCard TEXT value.</summary>
    public static string EncodeText(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        return value.Replace("\\", "\\\\").Replace(",", "\\,").Replace(";", "\\;")
            .Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n");
    }
}
