namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private const long MaximumOutputCharacters = 32 * 1024 * 1024;
    private int _resourceBindings;

    private void EnsureOutputCapacity(long characters) {
        _token.ThrowIfCancellationRequested();
        if (characters > MaximumOutputCharacters - _outputCharacters) throw new InvalidDataException("XPS SVG output budget exceeded.");
    }
    private void ChargeOutput(long characters) {
        EnsureOutputCapacity(characters);
        _outputCharacters += characters;
    }
    // Charge every expanded SVG node and attribute, including repeated resource projections,
    // before retaining it. The estimate includes XML escaping and namespace/element overhead.
    private XElement Element(string name, params object[] content) {
        _token.ThrowIfCancellationRequested();
        if (++_outputNodes > 100000) throw new InvalidDataException("XPS SVG node budget exceeded.");
        long size = 128 + name.Length * 2L;
        foreach (object item in content) if (item is XAttribute attribute) size += AttributeSize(attribute.Name.LocalName, attribute.Value);
        ChargeOutput(size);
        return new XElement(Svg + name, content);
    }
    private void Set(XElement element, string name, object value) {
        string text = System.Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        ChargeOutput(AttributeSize(name, text));
        element.SetAttributeValue(name, text);
    }
    private static long AttributeSize(string name, string value) {
        long length = name.Length + value.Length + 4L;
        foreach (char c in value) {
            if (c == '&') length += 4;
            else if (c == '"') length += 5;
            else if (c == '<' || c == '>') length += 3;
            else if (c == '\r' || c == '\n' || c == '\t') length += 5;
        }
        return length;
    }
    private void ChargeBindings(int count) {
        if (count > 100000 - _resourceBindings) throw new InvalidDataException("XPS resource binding budget exceeded.");
        _resourceBindings += count;
    }
}
