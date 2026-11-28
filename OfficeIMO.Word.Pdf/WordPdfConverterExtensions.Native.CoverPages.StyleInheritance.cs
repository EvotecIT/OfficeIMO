
namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private const int MaxNativeVmlStyleExtensionAttributes = 256;

        private static string? GetNativeVmlStyleValue(OfficeXmlNode element, OfficeXmlNode shapeType,
            string childName, string childAttribute, string shapeAttribute) {
            // Children override their shape's shorthand. Instance settings then
            // override the corresponding inherited shape-type settings.
            return GetNativeVmlOptionalStyleAttribute(GetNativeVmlChild(element, childName), childAttribute)
                ?? GetNativeOpenXmlAttribute(element, shapeAttribute)
                ?? GetNativeVmlOptionalStyleAttribute(GetNativeVmlChild(shapeType, childName), childAttribute)
                ?? GetNativeVmlOptionalStyleAttribute(shapeType, shapeAttribute);
        }

        private static string? GetNativeVmlOptionalStyleAttribute(OfficeXmlNode element, string attribute) =>
            element.Exists ? GetNativeOpenXmlAttribute(element, attribute) : null;

        private static OfficeXmlNode GetNativeVmlEffectiveStyleChild(OfficeXmlNode element,
            OfficeXmlNode shapeType, string childName) {
            OfficeXmlNode child = GetNativeVmlChild(element, childName);
            OfficeXmlNode inherited = GetNativeVmlChild(shapeType, childName);
            if (inherited.Exists && CountNativeVmlExtensionAttributes(inherited) > MaxNativeVmlStyleExtensionAttributes)
                throw new NotSupportedException("VML fill and stroke styles support at most 256 extension attributes.");
            if (child.IsNone) return inherited;
            if (inherited.IsNone) return child;

            // Merge on a detached copy so partial overrides retain inherited
            // opacity, gradients and line settings without changing the DOCX.
            OfficeXmlNode effective = child.Clone();
            var existingAttributes = new HashSet<(string LocalName, string NamespaceUri)>(
                effective.Attributes().Select(attribute => (attribute.LocalName, attribute.NamespaceUri)));
            foreach (OfficeXmlAttribute attribute in inherited.Attributes()) {
                if (IsNativeVmlRenderedStyleAttribute(childName, attribute.LocalName) &&
                    existingAttributes.Add((attribute.LocalName, attribute.NamespaceUri))) {
                    if (attribute.NamespaceUri.Length == 0) effective.SetAttribute(attribute.LocalName, attribute.Value);
                    else if (GetNativeVmlStyleNamespaceId(attribute.NamespaceUri) is int namespaceId)
                        effective.SetAttribute(namespaceId, attribute.LocalName, attribute.Value);
                }
            }
            return effective;
        }

        private static int? GetNativeVmlStyleNamespaceId(string namespaceUri) => namespaceUri switch {
            "urn:schemas-microsoft-com:vml" => WordXml.V,
            "urn:schemas-microsoft-com:office:office" => WordXml.O,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships" => WordXml.R,
            _ => null
        };

        // Extension attributes are those outside the VML, Office and relationship vocabularies the style elements declare.
        private static int CountNativeVmlExtensionAttributes(OfficeXmlNode element) =>
            element.Attributes().Count(attribute => attribute.NamespaceUri.Length == 0
                ? !NativeVmlStyleSchemaAttributes.Contains(attribute.LocalName)
                : attribute.NamespaceUri != "urn:schemas-microsoft-com:office:office" &&
                  attribute.NamespaceUri != "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // The unqualified attributes ECMA-376 Part 4 declares on v:fill and v:stroke.
        private static readonly HashSet<string> NativeVmlStyleSchemaAttributes = new(StringComparer.Ordinal) {
            "id", "type", "on", "color", "opacity", "color2", "src", "size", "origin", "position", "aspect", "colors",
            "angle", "alignshape", "focus", "focussize", "focusposition", "method", "detectmouseclick", "title", "recolor",
            "rotate", "weight", "linestyle", "miterlimit", "joinstyle", "endcap", "dashstyle", "filltype", "imageaspect",
            "imagesize", "imagealignshape", "startarrow", "startarrowwidth", "startarrowlength", "endarrow",
            "endarrowwidth", "endarrowlength", "insetpen"
        };

        private static bool IsNativeVmlRenderedStyleAttribute(string childName, string attributeName) =>
            childName == "fill"
                ? attributeName.Equals("type", StringComparison.OrdinalIgnoreCase)
                    || attributeName.Equals("color2", StringComparison.OrdinalIgnoreCase)
                    || attributeName.Equals("angle", StringComparison.OrdinalIgnoreCase)
                    || attributeName.Equals("colors", StringComparison.OrdinalIgnoreCase)
                    || attributeName.Equals("opacity", StringComparison.OrdinalIgnoreCase)
                : attributeName.Equals("dashstyle", StringComparison.OrdinalIgnoreCase)
                    || attributeName.Equals("endcap", StringComparison.OrdinalIgnoreCase)
                    || attributeName.Equals("joinstyle", StringComparison.OrdinalIgnoreCase)
                    || attributeName.Equals("opacity", StringComparison.OrdinalIgnoreCase);
    }
}
