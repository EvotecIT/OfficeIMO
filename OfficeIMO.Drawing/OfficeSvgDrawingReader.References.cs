using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumElementReferenceDepth = 16;

    private static void AddReferencedElement(
        XElement use,
        OfficeDrawing drawing,
        SvgPaintContext style,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform transform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref int unsupported) {
        if (!references.TryEnter(use, out string referenceId, out XElement? target)) {
            unsupported++;
            return;
        }

        try {
            string targetName = target!.Name.LocalName.ToLowerInvariant();
            if (targetName is "defs" or "lineargradient" or "radialgradient" or "stop") {
                unsupported++;
                return;
            }
            if (targetName == "symbol") {
                AddReferencedSymbol(use, target, drawing, style, paintServers, references, transform,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                    ref visited, ref pathCommands, ref unsupported);
                return;
            }
            if (!TryOptionalUseLength(use, "x", out double x) || !TryOptionalUseLength(use, "y", out double y)) {
                unsupported++;
                return;
            }

            OfficeTransform placement = OfficeTransform.Translate(x, y).Then(transform);
            AddElement(target, drawing, style, paintServers, references, placement, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                ref visited, ref pathCommands, ref unsupported);
        } finally {
            references.Exit(referenceId);
        }
    }

    private static bool TryOptionalUseLength(XElement use, string name, out double value) {
        string? text = use.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(text)) {
            value = 0D;
            return true;
        }
        return TrySvgLength(text, out value);
    }

    private sealed class SvgElementReferenceRegistry {
        private readonly SvgDefinitionRegistry _definitions;
        private readonly ISet<string> _activeIds = new HashSet<string>(StringComparer.Ordinal);

        internal SvgElementReferenceRegistry(SvgDefinitionRegistry definitions) {
            _definitions = definitions;
        }

        internal bool TryEnter(XElement use, out string id, out XElement? target) {
            id = string.Empty;
            target = null;
            XAttribute? href = use.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName.Equals("href", StringComparison.OrdinalIgnoreCase));
            if (href == null
                || !TryReadLocalElementReference(href.Value, out id)
                || !_definitions.TryGetUnique(id, out target)
                || _activeIds.Count >= MaximumElementReferenceDepth
                || !_activeIds.Add(id)) return false;
            return true;
        }

        internal void Exit(string id) {
            if (id.Length > 0) _activeIds.Remove(id);
        }

        private static bool TryReadLocalElementReference(string text, out string id) {
            id = string.Empty;
            string normalized = text.Trim();
            if (normalized.Length < 2 || normalized[0] != '#') return false;
            id = normalized.Substring(1);
            return id.Length > 0 && id.IndexOfAny(new[] { ' ', '\t', '\r', '\n', '#', '(', ')' }) < 0;
        }
    }
}
