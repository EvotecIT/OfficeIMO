using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const int MaxNativeVmlStyleExtensionAttributes = 256;

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
            if (inherited != null && inherited.ExtendedAttributes.Skip(MaxNativeVmlStyleExtensionAttributes).Any())
                throw new NotSupportedException("VML fill and stroke styles support at most 256 extension attributes.");
            if (child == null) return inherited;
            if (inherited == null) return child;

            // Merge on a detached copy so partial overrides retain inherited
            // opacity, gradients and line settings without changing the DOCX.
            OpenXmlElement effective = child.CloneNode(true);
            var existingAttributes = new HashSet<(string LocalName, string NamespaceUri)>(
                effective.GetAttributes().Select(attribute => (attribute.LocalName, attribute.NamespaceUri)));
            foreach (OpenXmlAttribute attribute in inherited.GetAttributes()) {
                if (IsNativeVmlRenderedStyleAttribute(childName, attribute.LocalName) &&
                    existingAttributes.Add((attribute.LocalName, attribute.NamespaceUri))) {
                    effective.SetAttribute(attribute);
                }
            }
            return effective;
        }

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
