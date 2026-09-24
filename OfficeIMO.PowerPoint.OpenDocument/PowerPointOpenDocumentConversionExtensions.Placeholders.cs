using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static string? GetOdpPresentationClass(PowerPointSlide slide, PowerPointTextBox textBox) {
        if (!textBox.IsPlaceholder) return null;
        PowerPointPlaceholderType? type = textBox.PlaceholderType;
        if (!type.HasValue && textBox.PlaceholderIndex.HasValue) {
            type = slide.GetLayoutPlaceholders().FirstOrDefault(placeholder =>
                placeholder.PlaceholderIndex == textBox.PlaceholderIndex).PlaceholderType;
        }
        return type switch {
            PowerPointPlaceholderType.Title or PowerPointPlaceholderType.CenteredTitle => "title",
            PowerPointPlaceholderType.SubTitle => "subtitle",
            PowerPointPlaceholderType.Body => "outline",
            PowerPointPlaceholderType.Object => "object",
            _ => null
        };
    }

    private static PowerPointPlaceholderType? GetPowerPointPlaceholderType(string? presentationClass) =>
        presentationClass switch {
            "title" => PowerPointPlaceholderType.Title,
            "subtitle" => PowerPointPlaceholderType.SubTitle,
            "outline" or "text" => PowerPointPlaceholderType.Body,
            "object" => PowerPointPlaceholderType.Object,
            _ => null
        };
}
