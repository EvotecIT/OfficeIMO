using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private sealed class SvgRenderContext {
            private readonly Dictionary<string, XElement> _definitions;
            private readonly HashSet<string> _activeUseIds = new(StringComparer.Ordinal);

            private readonly Func<string, byte[]?>? _imageResolver;

            private SvgRenderContext(SvgStyleSheet styleSheet, Dictionary<string, XElement> definitions, Func<string, byte[]?>? imageResolver) {
                StyleSheet = styleSheet;
                _definitions = definitions;
                _imageResolver = imageResolver;
            }

            internal SvgStyleSheet StyleSheet { get; }

            internal static SvgRenderContext Create(XElement root, Func<string, byte[]?>? imageResolver = null) =>
                new(SvgStyleSheet.Parse(root), ReadDefinitions(root), imageResolver);

            internal bool TryGetDefinition(string id, out XElement? definition) =>
                _definitions.TryGetValue(id, out definition);

            internal bool TryEnterUse(string id) => _activeUseIds.Add(id);

            internal void ExitUse(string id) => _activeUseIds.Remove(id);

            internal bool TryGetImageBytes(string href, out byte[]? bytes) {
                bytes = _imageResolver?.Invoke(href);
                return bytes != null && bytes.Length > 0;
            }

            private static Dictionary<string, XElement> ReadDefinitions(XElement root) {
                Dictionary<string, XElement> definitions = new(StringComparer.Ordinal);
                foreach (XElement element in root.Descendants()) {
                    string? id = element.Attribute("id")?.Value;
                    if (!string.IsNullOrWhiteSpace(id) && !definitions.ContainsKey(id!)) {
                        definitions[id!] = element;
                    }
                }

                return definitions;
            }
        }
    }
}
