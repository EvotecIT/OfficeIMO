namespace OfficeIMO.Email;

/// <summary>Projects MAPI address properties without losing Exchange or source-specific values.</summary>
internal static class MsgAddressProjection {
    internal static EmailAddress? ReadAddress(
        IEnumerable<MapiProperty> properties,
        MapiPropertyKey<string> displayNameKey,
        MapiPropertyKey<string> smtpAddressKey,
        MapiPropertyKey<string> emailAddressKey,
        MapiPropertyKey<string> addressTypeKey,
        MapiPropertyKey<string>? originalAddressKey = null) {
        string? displayName = Clean(properties.GetMapiValueOrDefault(displayNameKey));
        string? addressType = Clean(properties.GetMapiValueOrDefault(addressTypeKey));
        string? smtp = Clean(properties.GetMapiValueOrDefault(smtpAddressKey));
        string? native = Clean(properties.GetMapiValueOrDefault(emailAddressKey));
        string? original = originalAddressKey != null
            ? Clean(properties.GetMapiValueOrDefault(originalAddressKey))
            : null;

        string? address = string.Equals(addressType, "EX", StringComparison.OrdinalIgnoreCase)
            ? FirstNonEmpty(smtp, native)
            : FirstNonEmpty(native, smtp);
        if ((string.IsNullOrEmpty(address) || (address?.IndexOf('@') ?? -1) < 0) &&
            !string.IsNullOrEmpty(original) && (original?.IndexOf('@') ?? -1) >= 0) {
            address = original;
        }

        if (!LooksLikeInternetAddress(address) && LooksLikeInternetAddress(displayName)) {
            string? swap = address;
            address = displayName;
            displayName = swap;
        }
        if (string.Equals(address, displayName, StringComparison.OrdinalIgnoreCase)) displayName = null;
        if (address == null && displayName == null && native == null) return null;

        return new EmailAddress(address, displayName, native ?? original) {
            AddressType = addressType
        };
    }

    internal static EmailRecipientKind ReadRecipientKind(IEnumerable<MapiProperty> properties) {
        int displayTypeEx = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.DisplayTypeEx) ?? -1;
        if ((displayTypeEx & 0xFF) == 7) return EmailRecipientKind.Room;

        return (properties.GetNullableMapiValue(MapiKnownProperties.PidTag.RecipientType) ?? 0) switch {
            1 => EmailRecipientKind.To,
            2 => EmailRecipientKind.Cc,
            3 => EmailRecipientKind.Bcc,
            4 => EmailRecipientKind.Resource,
            _ => EmailRecipientKind.Unknown
        };
    }

    private static string? Clean(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        return value!.Trim().Trim('\'', '"');
    }

    private static string? FirstNonEmpty(string? first, string? second) =>
        !string.IsNullOrEmpty(first) ? first : second;

    private static bool LooksLikeInternetAddress(string? value) =>
        !string.IsNullOrEmpty(value) && value!.IndexOf('@') > 0;
}
