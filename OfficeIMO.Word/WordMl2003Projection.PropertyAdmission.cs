using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private OpenXmlValidator? _propertyValidator;

    private void AdmitMappedProperties(OpenXmlCompositeElement properties) {
        if (!properties.HasChildren) return;

        // Match WordDocument.ValidateDocument's default profile and reuse the SDK cache.
        OpenXmlValidator validator = _propertyValidator ??= new OpenXmlValidator(FileFormatVersions.Microsoft365);
        ValidationErrorInfo[] errors = validator.Validate(properties, _token).ToArray();
        if (validator.MaxNumberOfErrors > 0 && errors.Length >= validator.MaxNumberOfErrors)
            throw new InvalidDataException("WordML 2003 property group exceeds the SDK validation error limit.");

        var invalid = new HashSet<OpenXmlElement>();
        foreach (ValidationErrorInfo error in errors) {
            OpenXmlElement? node = error.Node;
            // Wrapper attributes have separate readers or enum admission; their children are mapped properties.
            if (node == null || ReferenceEquals(node, properties) || !invalid.Add(node)) continue;

            string name = LegacyPropertyName(node.LocalName);
            string attributes = string.Join(", ", node.GetAttributes().Select(DescribeMappedAttribute));
            Omit(Word + name, name + " property with unsupported attribute values [" + attributes + "] (" + error.Id + ")");
        }

        foreach (OpenXmlElement node in invalid) {
            OpenXmlElement? ancestor = node.Parent;
            while (ancestor != null && !ReferenceEquals(ancestor, properties)) ancestor = ancestor.Parent;
            if (ancestor == null) continue; // An invalid ancestor has already been omitted.

            OpenXmlElement? parent = node.Parent;
            node.Remove();
            while (parent?.Parent != null && !ReferenceEquals(parent, properties) &&
                !parent.HasChildren && !parent.HasAttributes) {
                OpenXmlElement next = parent.Parent;
                parent.Remove();
                parent = next;
            }
        }
    }

    private static string DescribeMappedAttribute(OpenXmlAttribute attribute) {
        string name = attribute.LocalName;
        foreach (KeyValuePair<string, string> mapping in AttributeNames) {
            if (mapping.Value == name) { name = mapping.Key; break; }
        }
        string value = attribute.Value ?? string.Empty;
        if (value.Length > 80) value = value.Substring(0, 80) + "...";
        return name + "='" + value + "'";
    }
}
