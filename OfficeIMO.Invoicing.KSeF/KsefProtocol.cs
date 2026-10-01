using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeIMO.Invoicing.KSeF;

internal static class KsefProtocol {
    internal static string Reference(string value) {
        if (value == null || value.Length != 36 || !Regex.IsMatch(value, @"\A[0-9]{8}-[0-9A-Z]{2}-[0-9A-F]{10}-[0-9A-F]{10}-[0-9A-F]{2}\z")) throw new InvalidDataException("Invalid KSeF operation reference.");
        return value;
    }
    internal static string Challenge(string value) {
        if (value == null || value.Length != 36 || !Regex.IsMatch(value, @"\A[0-9]{8}-CR-[0-9A-F]{10}-[0-9A-F]{10}-[0-9A-F]{2}\z")) throw new InvalidDataException("Invalid KSeF challenge.");
        return value;
    }
    internal static string InvoiceNumber(string value) {
        if (value == null || value.Length is not (35 or 36) || !Regex.IsMatch(value, @"\A[1-9]([0-9][1-9]|[1-9][0-9])[0-9]{7}-(20[2-9][0-9]|2[1-9][0-9]{2}|[3-9][0-9]{3})(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])-[0-9A-F]{6}-?[0-9A-F]{6}-[0-9A-F]{2}\z")) throw new InvalidDataException("Invalid KSeF invoice number.");
        return value;
    }
    internal static string Hash(string value) {
        if (value == null || value.Length != 44) throw new InvalidDataException("Expected a Base64 SHA-256 digest.");
        Span<byte> bytes = stackalloc byte[32];
        if (!Convert.TryFromBase64String(value, bytes, out int count) || count != 32 || Convert.ToBase64String(bytes) != value) throw new InvalidDataException("Expected a canonical Base64 SHA-256 digest.");
        return value;
    }
    internal static string Text(JsonElement parent, string name, int maximum = 256) {
        if (!parent.TryGetProperty(name, out JsonElement element) || element.ValueKind != JsonValueKind.String) throw new InvalidDataException("KSeF response lacks a required string field.");
        string value = element.GetString()!;
        if (value.Length == 0 || value.Length > maximum || value.Any(char.IsControl)) throw new InvalidDataException("KSeF response field exceeds its supported bounds.");
        return value;
    }
    internal static DateTimeOffset Instant(JsonElement parent, string name) {
        string value = Text(parent, name, 64);
        if (!(value.EndsWith('Z') || Regex.IsMatch(value, @"[+-][0-9]{2}:[0-9]{2}\z")) || !DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTimeOffset result)) throw new InvalidDataException("KSeF response timestamp must include a timezone.");
        return result;
    }
    internal static KsefStatus Status(JsonElement parent) {
        JsonElement status = parent.GetProperty("status");
        return new KsefStatus(status.GetProperty("code").GetInt32(), Text(status, "description", 4096));
    }
    internal static byte[] Json(Action<Utf8JsonWriter> write) {
        using var output = new MemoryStream(); using (var writer = new Utf8JsonWriter(output)) { writer.WriteStartObject(); write(writer); writer.WriteEndObject(); }
        return output.ToArray();
    }
}
