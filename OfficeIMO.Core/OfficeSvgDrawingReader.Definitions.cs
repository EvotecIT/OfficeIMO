using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed class SvgDefinitionRegistry {
        private readonly IReadOnlyDictionary<string, XElement> _definitions;
        private readonly ISet<string> _ambiguousIds;

        private SvgDefinitionRegistry(IReadOnlyDictionary<string, XElement> definitions, ISet<string> ambiguousIds) {
            _definitions = definitions;
            _ambiguousIds = ambiguousIds;
        }

        internal static SvgDefinitionRegistry Create(XElement root) {
            var definitions = new Dictionary<string, XElement>(StringComparer.Ordinal);
            var ambiguousIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement element in root.Descendants()) {
                string? id = element.Attribute("id")?.Value.Trim();
                if (string.IsNullOrEmpty(id)) continue;
                if (definitions.ContainsKey(id!)) {
                    ambiguousIds.Add(id!);
                    continue;
                }
                definitions.Add(id!, element);
            }
            return new SvgDefinitionRegistry(definitions, ambiguousIds);
        }

        internal bool TryGetUnique(string id, out XElement? element) {
            element = null;
            return !_ambiguousIds.Contains(id) && _definitions.TryGetValue(id, out element);
        }
    }
}
