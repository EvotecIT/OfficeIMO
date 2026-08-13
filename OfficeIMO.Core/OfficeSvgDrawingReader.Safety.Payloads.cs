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
}
