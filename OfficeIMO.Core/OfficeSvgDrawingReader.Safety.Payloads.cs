using System;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int RenderedSvgPayloadCharactersPerElement = 100;

    private static bool TryAddRenderedSvgPayloadComplexity(
        XElement element,
        int maximumElements,
        ref int elementCount) {
        int directCharacters = 0;
        foreach (XAttribute attribute in element.Attributes()) {
            if (!TryAddRenderedSvgPayloadCharacters(attribute.Value.Length, maximumElements, ref directCharacters)) return false;
        }
        foreach (XNode node in element.Nodes()) {
            if (node is XText text
                && !TryAddRenderedSvgPayloadCharacters(text.Value.Length, maximumElements, ref directCharacters)) return false;
        }
        if (element.Name.LocalName.Equals("image", StringComparison.OrdinalIgnoreCase)
            && !IsWithinEmbeddedRasterImageLimits(element)) return false;
        if (directCharacters <= RenderedSvgPayloadCharactersPerElement) return true;

        int additionalElements = (directCharacters - 1) / RenderedSvgPayloadCharactersPerElement;
        if (additionalElements > maximumElements - elementCount) return false;
        elementCount += additionalElements;
        return true;
    }

    private static bool TryAddRenderedSvgPayloadCharacters(
        int characters,
        int maximumElements,
        ref int directCharacters) {
        if (characters > maximumElements * RenderedSvgPayloadCharactersPerElement - directCharacters) return false;
        directCharacters += characters;
        return true;
    }

    private static bool IsWithinEmbeddedRasterImageLimits(XElement element) {
        string? href = ReadRasterProjectedAttribute(element, "href")?.Trim();
        if (string.IsNullOrEmpty(href) || !href!.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return true;
        int comma = href.IndexOf(',');
        if (comma <= 5 || comma == href.Length - 1) return false;
        string metadata = href.Substring(5, comma - 5);
        int separator = metadata.IndexOf(';');
        string mediaType = (separator < 0 ? metadata : metadata.Substring(0, separator)).Trim();
        if (!mediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
            || !ContainsBase64DataUriToken(metadata)) return false;

        byte[] imageBytes;
        try {
            imageBytes = Convert.FromBase64String(href.Substring(comma + 1));
        } catch (FormatException) {
            return false;
        }
        return OfficeImageReader.TryIdentifyByContent(imageBytes, null, out OfficeImageInfo info)
            && info.Format != OfficeImageFormat.Svg
            && OfficeRasterGuards.TryEnsurePixelCount(info.Width, info.Height, out _);
    }

    private static bool ContainsBase64DataUriToken(string metadata) {
        foreach (string token in metadata.Split(';')) {
            if (token.Trim().Equals("base64", StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }
}
