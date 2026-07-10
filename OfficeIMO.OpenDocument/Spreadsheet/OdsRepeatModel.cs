namespace OfficeIMO.OpenDocument;

internal static class OdsRepeatModel {
    internal static long Read(XElement element, XName attribute) {
        string? lexical = (string?)element.Attribute(attribute);
        if (lexical == null) return 1;
        if (!long.TryParse(lexical, NumberStyles.None, CultureInfo.InvariantCulture, out long value) || value < 1) {
            throw new InvalidDataException($"Invalid ODF repeat count '{lexical}' on '{element.Name.LocalName}'.");
        }
        return value;
    }

    internal static void Set(XElement element, XName attribute, long value) {
        if (value < 1) throw new ArgumentOutOfRangeException(nameof(value));
        element.SetAttributeValue(attribute, value == 1 ? (long?)null : value);
    }

    internal static XElement Split(XElement element, XName repeatAttribute, long offset) {
        long count = Read(element, repeatAttribute);
        if (offset < 0 || offset >= count) throw new ArgumentOutOfRangeException(nameof(offset));
        if (count == 1) return element;

        var replacements = new List<XElement>(3);
        if (offset > 0) {
            XElement before = new XElement(element);
            Set(before, repeatAttribute, offset);
            replacements.Add(before);
        }
        XElement target = new XElement(element);
        Set(target, repeatAttribute, 1);
        replacements.Add(target);
        long afterCount = count - offset - 1;
        if (afterCount > 0) {
            XElement after = new XElement(element);
            Set(after, repeatAttribute, afterCount);
            replacements.Add(after);
        }
        element.ReplaceWith(replacements);
        return target;
    }
}
