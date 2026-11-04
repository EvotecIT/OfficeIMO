namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    /// <summary>Admits the ordinary text-property family observed in the native controls, without interpreting unqualified direction, anchor or fitting overrides.</summary>
    private static string? ImportedFrameLayoutLimitation(PublisherSourceShape native, PublisherEscherShape? shape, uint vertical) {
        PublisherBlock? verticalField = PublisherContentsReader.Field(native.Fields, 0x35);
        if (verticalField.HasValue && verticalField.Value.Type != 0x20)
            return "The native vertical-alignment encoding is not qualified for layout editing.";
        if (vertical > 2) return "The native vertical alignment is outside the qualified Top, Center and Bottom family.";
        if (shape == null) return "The frame has no recovered drawing layout.";
        foreach (var property in shape.Properties.Where(item => item.PropertyId is >= 0x80 and <= 0xBF)) {
            if (property.IsComplex || property.IsBlipId) return "A native text-layout property uses unqualified complex or image-reference flags.";
            switch (property.PropertyId) {
                case 0x80: // Native story binding; current story ownership is separate.
                case 0x81: case 0x82: case 0x83: case 0x84:
                case 0x8C: case 0x8D:
                    break;
                case 0x85 when property.Value == 1:
                case 0xBF when property.Value == 0x00020000:
                    // Exact disabled/default forms present in independently
                    // produced Text/Newsletter/master controls. Other values
                    // retain reading but cannot be silently treated as defaults.
                    break;
                default:
                    return "A native text wrapping, anchor, direction or fitting override is not qualified for layout editing.";
            }
        }
        return null;
    }
}
