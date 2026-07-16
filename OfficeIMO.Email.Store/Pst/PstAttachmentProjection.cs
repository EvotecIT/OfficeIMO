using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class PstAttachmentProjection {
    internal static EmailAttachment Create(IReadOnlyList<MapiProperty> properties,
        EmailStoreReaderOptions options, ref long totalAttachmentBytes) {
        byte[]? content = GetBytes(properties, 0x3701);
        long length = content?.LongLength ?? 0;
        if (length > options.MaxAttachmentBytes) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                length, options.MaxAttachmentBytes);
        }
        totalAttachmentBytes = checked(totalAttachmentBytes + length);
        if (totalAttachmentBytes > options.MaxTotalAttachmentBytes) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes),
                totalAttachmentBytes, options.MaxTotalAttachmentBytes);
        }

        var attachment = new EmailAttachment {
            FileName = GetString(properties, 0x3707) ?? GetString(properties, 0x3704) ?? GetString(properties, 0x3001),
            ContentType = GetString(properties, 0x370E),
            ContentId = TrimAngle(GetString(properties, 0x3712)),
            ContentLocation = GetString(properties, 0x3713),
            RenderingPosition = GetInt(properties, 0x370B) ?? -1,
            IsInline = (GetInt(properties, 0x370B) ?? -1) >= 0,
            IsHidden = GetBool(properties, 0x7FFE) ?? false,
            CreatedDate = GetDate(properties, 0x3007),
            ModifiedDate = GetDate(properties, 0x3008),
            LinkedPath = GetString(properties, 0x370D),
            Length = length,
            Content = options.RetainAttachmentContent ? content : null,
            MapiAttachMethod = GetInt(properties, 0x3705)
        };
        foreach (MapiProperty property in properties) attachment.MapiProperties.Add(property);
        return attachment;
    }

    private static string? GetString(IEnumerable<MapiProperty> properties, ushort id) =>
        properties.FirstOrDefault(property => property.PropertyId == id)?.Value as string;

    private static byte[]? GetBytes(IEnumerable<MapiProperty> properties, ushort id) =>
        properties.FirstOrDefault(property => property.PropertyId == id)?.Value as byte[];

    private static int? GetInt(IEnumerable<MapiProperty> properties, ushort id) {
        object? value = properties.FirstOrDefault(property => property.PropertyId == id)?.Value;
        if (value is int number) return number;
        if (value is short shortNumber) return shortNumber;
        return null;
    }

    private static bool? GetBool(IEnumerable<MapiProperty> properties, ushort id) =>
        properties.FirstOrDefault(property => property.PropertyId == id)?.Value as bool?;

    private static DateTimeOffset? GetDate(IEnumerable<MapiProperty> properties, ushort id) =>
        properties.FirstOrDefault(property => property.PropertyId == id)?.Value as DateTimeOffset?;

    private static string? TrimAngle(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return value;
        string trimmed = value!.Trim();
        return trimmed.Length > 1 && trimmed[0] == '<' && trimmed[trimmed.Length - 1] == '>'
            ? trimmed.Substring(1, trimmed.Length - 2)
            : trimmed;
    }
}
