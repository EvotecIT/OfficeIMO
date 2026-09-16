using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed class SvgDefinitionRegistry {
        private readonly IReadOnlyDictionary<string, XElement> _definitions;
        private readonly ISet<string> _ambiguousIds;

        private SvgDefinitionRegistry(
            XNamespace nativeNamespace,
            IReadOnlyDictionary<string, XElement> definitions,
            ISet<string> ambiguousIds) {
            NativeNamespace = nativeNamespace;
            _definitions = definitions;
            _ambiguousIds = ambiguousIds;
        }

        internal XNamespace NativeNamespace { get; }

        internal static SvgDefinitionRegistry Create(XElement root, bool useProjectedIds = false) {
            var definitions = new Dictionary<string, XElement>(StringComparer.Ordinal);
            var ambiguousIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement element in root.Descendants()) {
                if (!IsNativeSvgElement(element, root.Name.Namespace)) continue;
                string? id = useProjectedIds
                    ? ReadRasterProjectedAttribute(element, "id")
                    : ReadRasterElementId(element);
                if (string.IsNullOrEmpty(id)) continue;
                if (definitions.ContainsKey(id!)) {
                    ambiguousIds.Add(id!);
                    continue;
                }
                definitions.Add(id!, element);
            }
            return new SvgDefinitionRegistry(root.Name.Namespace, definitions, ambiguousIds);
        }

        internal bool TryGetUnique(string id, out XElement? element) {
            element = null;
            return !_ambiguousIds.Contains(id) && _definitions.TryGetValue(id, out element);
        }
    }

    private static string? ReadRasterElementId(XElement element) {
        string? id = element.Attribute("id")?.Value;
        return string.IsNullOrEmpty(id) ? null : id;
    }

    private static string? ReadRasterProjectedAttribute(XElement element, string localName) {
        string? value = null;
        foreach (XAttribute attribute in element.Attributes()) {
            if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == XNamespace.Xml) continue;
            if (attribute.Name.LocalName.Equals(localName, StringComparison.Ordinal)) value = attribute.Value;
        }
        return value;
    }
}
