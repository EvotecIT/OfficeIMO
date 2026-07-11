namespace OfficeIMO.Email;

/// <summary>Projects MAPI address properties without losing Exchange or source-specific values.</summary>
internal static class MsgAddressProjection {
    internal static EmailAddress? ReadAddress(
        IEnumerable<MapiProperty> properties,
        ushort displayNameId,
        ushort smtpAddressId,
        ushort emailAddressId,
        ushort addressTypeId,
        ushort? originalAddressId = null) {
        string? displayName = Clean(MsgProjection.GetString(properties, displayNameId));
        string? addressType = Clean(MsgProjection.GetString(properties, addressTypeId));
        string? smtp = Clean(MsgProjection.GetString(properties, smtpAddressId));
        string? native = Clean(MsgProjection.GetString(properties, emailAddressId));
        string? original = originalAddressId.HasValue
            ? Clean(MsgProjection.GetString(properties, originalAddressId.Value))
            : null;

        string? address = string.Equals(addressType, "EX", StringComparison.OrdinalIgnoreCase)
            ? FirstNonEmpty(smtp, native)
            : FirstNonEmpty(native, smtp);
        if ((string.IsNullOrEmpty(address) || address.IndexOf('@') < 0) &&
            !string.IsNullOrEmpty(original) && original.IndexOf('@') >= 0) {
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
        int displayTypeEx = MsgProjection.GetInt(properties, 0x3905) ?? -1;
        if ((displayTypeEx & 0xFF) == 7) return EmailRecipientKind.Room;

        return (MsgProjection.GetInt(properties, 0x0C15) ?? 0) switch {
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
