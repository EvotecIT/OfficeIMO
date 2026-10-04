namespace OfficeIMO.Xps;

// Native URI-bearing color syntax is shared by rendering and package relationship generation.
internal static class XpsResourceSyntax {
    internal static bool IsContextColor(string value) => value.StartsWith("ContextColor", StringComparison.Ordinal);

    internal static (string Uri, double[] Values) ContextColor(string value) {
        string[] fields = value.Split(new[] { ' ', '\t', '\r', '\n' }, 3, StringSplitOptions.RemoveEmptyEntries);
        if (fields.Length != 3 || fields[0] != "ContextColor") throw new InvalidDataException("Invalid ContextColor syntax.");
        string[] samples = fields[2].Split(',');
        if (samples.Length < 2 || samples.Length > 9) throw new InvalidDataException("Invalid ContextColor channel count.");
        return (fields[1], samples.Select(s => XpsPackage.Number(s.Trim())).ToArray());
    }

    internal static IEnumerable<string> References(XElement markup) {
        foreach (var attribute in markup.DescendantsAndSelf().Attributes()) {
            string value = attribute.Value;
            if (attribute.Name.LocalName == "FontUri" || attribute.Name.LocalName == "ImageSource" ||
                (attribute.Name.LocalName == "Source" && attribute.Parent?.Name.LocalName == "ResourceDictionary")) {
                if (!value.StartsWith("{", StringComparison.Ordinal)) yield return value.Split('#')[0];
            } else if ((attribute.Name.LocalName == "Color" || attribute.Name.LocalName == "Fill" ||
                attribute.Name.LocalName == "Stroke" || attribute.Name.LocalName == "OpacityMask") && IsContextColor(value)) {
                yield return ContextColor(value).Uri;
            }
        }
    }
}
