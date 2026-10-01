using System.Net.Mail;

namespace OfficeIMO.Email;

/// <summary>Shared SMTP projection for independent drafts and field-selected copies.</summary>
internal static class EmailSmtpAddress {
    internal static string? Normalize(EmailAddress? address) {
        if (string.IsNullOrWhiteSpace(address?.Address) || address!.Address!.Any(char.IsControl) ||
            (!string.IsNullOrWhiteSpace(address.AddressType) && !string.Equals(address.AddressType, "SMTP", StringComparison.OrdinalIgnoreCase))) return null;
        try {
            string value = address.Address!;
            var parsed = new MailAddress(value);
            return parsed.Address.Contains("@") && string.Equals(parsed.Address, value.Trim(), StringComparison.OrdinalIgnoreCase) ? parsed.Address : null;
        } catch (FormatException) { return null; }
    }
}
