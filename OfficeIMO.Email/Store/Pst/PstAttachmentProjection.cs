using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class PstAttachmentProjection {
    internal static EmailAttachment Create(IReadOnlyList<MapiProperty> properties,
        long length, byte[]? content = null, IEmailContentSource? contentSource = null) {
        var attachment = new EmailAttachment {
            FileName = properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AttachLongFilename) ??
                properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AttachFilename) ??
                properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.DisplayName),
            ContentType = properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AttachMimeTag),
            ContentId = TrimAngle(properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AttachContentId)),
            ContentLocation = properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AttachContentLocation),
            RenderingPosition = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.RenderingPosition) ?? -1,
            IsInline = (properties.GetNullableMapiValue(MapiKnownProperties.PidTag.RenderingPosition) ?? -1) >= 0,
            IsHidden = properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AttachmentHidden),
            CreatedDate = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.CreationTime),
            ModifiedDate = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.LastModificationTime),
            LinkedPath = properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AttachLongPathname),
            Length = length,
            Content = content,
            ContentSource = contentSource,
            MapiAttachMethod = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.AttachMethod)
        };
        foreach (MapiProperty property in properties) attachment.MapiProperties.Add(property);
        return attachment;
    }

    private static string? TrimAngle(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return value;
        string trimmed = value!.Trim();
        return trimmed.Length > 1 && trimmed[0] == '<' && trimmed[trimmed.Length - 1] == '>'
            ? trimmed.Substring(1, trimmed.Length - 2)
            : trimmed;
    }
}
