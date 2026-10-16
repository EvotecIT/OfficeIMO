using DocumentFormat.OpenXml;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static string? GetNativeVmlStyleValue(OpenXmlElement element, OpenXmlElement? shapeType,
            string childName, string childAttribute, string shapeAttribute) {
            // Children override their shape's shorthand. Instance settings then
            // override the corresponding inherited shape-type settings.
            return GetNativeVmlOptionalStyleAttribute(GetNativeVmlChild(element, childName), childAttribute)
                ?? GetNativeOpenXmlAttribute(element, shapeAttribute)
                ?? GetNativeVmlOptionalStyleAttribute(GetNativeVmlChild(shapeType, childName), childAttribute)
                ?? GetNativeVmlOptionalStyleAttribute(shapeType, shapeAttribute);
        }

        private static string? GetNativeVmlOptionalStyleAttribute(OpenXmlElement? element, string attribute) =>
            element is not null ? GetNativeOpenXmlAttribute(element, attribute) : null;

        private static OpenXmlElement? GetNativeVmlEffectiveStyleChild(OpenXmlElement element,
            OpenXmlElement? shapeType, string childName) {
            OpenXmlElement? child = GetNativeVmlChild(element, childName);
            OpenXmlElement? inherited = GetNativeVmlChild(shapeType, childName);
            if (child == null) return inherited;
            if (inherited == null) return child;

            // Merge on a detached copy so partial overrides retain inherited
            // opacity, gradients and line settings without changing the DOCX.
            OpenXmlElement effective = child.CloneNode(true);
            foreach (OpenXmlAttribute attribute in inherited.GetAttributes()) {
                if (!effective.GetAttributes().Any(existing =>
                        existing.LocalName == attribute.LocalName && existing.NamespaceUri == attribute.NamespaceUri)) {
                    effective.SetAttribute(attribute);
                }
            }
            return effective;
        }
    }
}
