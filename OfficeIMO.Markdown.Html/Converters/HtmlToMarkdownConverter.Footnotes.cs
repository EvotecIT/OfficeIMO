using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    internal sealed class HtmlFootnoteConversionState {
        private readonly IReadOnlyList<IElement> _orderedDefinitions;
        private readonly Dictionary<IElement, string> _labelsByDefinition;
        private readonly Dictionary<string, string> _labelsByTargetId;

        private HtmlFootnoteConversionState(
            IReadOnlyList<IElement> orderedDefinitions,
            Dictionary<IElement, string> labelsByDefinition,
            Dictionary<string, string> labelsByTargetId) {
            _orderedDefinitions = orderedDefinitions;
            _labelsByDefinition = labelsByDefinition;
            _labelsByTargetId = labelsByTargetId;
        }

        internal static HtmlFootnoteConversionState Empty { get; } = new HtmlFootnoteConversionState(
            Array.Empty<IElement>(),
            new Dictionary<IElement, string>(),
            new Dictionary<string, string>(StringComparer.Ordinal));

        internal static HtmlFootnoteConversionState Create(INode root) {
            if (root == null) return Empty;

            List<IElement> elements = EnumerateElements(root).ToList();
            var referencedIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (IElement element in elements) {
                if (!LooksLikeFootnoteReference(element)) continue;
                string? targetId = GetLocalFragmentId(element.GetAttribute("href"));
                if (!string.IsNullOrWhiteSpace(targetId)) referencedIds.Add(targetId!);
            }

            var definitions = new List<IElement>();
            foreach (IElement element in elements) {
                if (IsSemanticFootnoteDefinition(element)
                    || IsRecognizedIdDefinition(element, referencedIds)) {
                    definitions.Add(element);
                }
            }
            if (definitions.Count == 0) return Empty;

            var labelsByDefinition = new Dictionary<IElement, string>();
            var labelsByTargetId = new Dictionary<string, string>(StringComparer.Ordinal);
            var usedLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < definitions.Count; index++) {
                IElement definition = definitions[index];
                string? definitionId = definition.Id;
                string label = CreateUniqueLabel(definitionId, index + 1, usedLabels);
                labelsByDefinition[definition] = label;
                if (!string.IsNullOrWhiteSpace(definitionId) && !labelsByTargetId.ContainsKey(definitionId!)) {
                    labelsByTargetId[definitionId!] = label;
                }
            }

            return new HtmlFootnoteConversionState(definitions, labelsByDefinition, labelsByTargetId);
        }

        internal bool TryGetDefinitionLabel(IElement element, out string label) =>
            _labelsByDefinition.TryGetValue(element, out label!);

        internal bool TryGetReferenceLabel(IElement element, out string label) {
            label = string.Empty;
            if (!LooksLikeFootnoteReference(element)) return false;
            string? targetId = GetLocalFragmentId(element.GetAttribute("href"));
            return targetId != null && _labelsByTargetId.TryGetValue(targetId, out label!);
        }

        internal bool TryGetWrappedReferenceLabel(IElement element, out string label) {
            label = string.Empty;
            if (element == null || !element.TagName.Equals("SUP", StringComparison.OrdinalIgnoreCase)) return false;

            IElement[] references = element.QuerySelectorAll("a")
                .Where(reference => TryGetReferenceLabel(reference, out _))
                .ToArray();
            if (references.Length != 1 || !TryGetReferenceLabel(references[0], out label)) return false;

            return string.Equals(
                NormalizeComparisonText(element.TextContent),
                NormalizeComparisonText(references[0].TextContent),
                StringComparison.Ordinal);
        }

        internal bool IsBacklink(IElement element) {
            if (element == null || !element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)) return false;
            if (HtmlAccessibilitySemantics.HasEpubType(element, "backlink")
                || HtmlAccessibilitySemantics.HasRole(element, "doc-backlink")
                || element.HasAttribute("data-footnote-backref")
                || element.ClassList.Contains("footnote-backref")) {
                return true;
            }

            string? href = element.GetAttribute("href");
            if (string.IsNullOrWhiteSpace(href)
                || (!href!.StartsWith("#fnref:", StringComparison.OrdinalIgnoreCase)
                    && !href.StartsWith("#fnref-", StringComparison.OrdinalIgnoreCase))) {
                return false;
            }
            for (IElement? current = element.ParentElement; current != null; current = current.ParentElement) {
                if (_labelsByDefinition.ContainsKey(current)) return true;
            }
            return false;
        }

        internal bool ShouldConvertContainer(IElement element) {
            if (element == null || _orderedDefinitions.Count == 0) return false;
            if (IsFootnoteContainer(element)) return _orderedDefinitions.Any(definition => IsDescendantOf(definition, element));

            string tagName = element.TagName;
            if (!tagName.Equals("OL", StringComparison.OrdinalIgnoreCase)
                && !tagName.Equals("UL", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            IElement[] items = element.Children
                .Where(static child => child.TagName.Equals("LI", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            return items.Length > 0 && items.All(item => _labelsByDefinition.ContainsKey(item));
        }

        internal IEnumerable<IElement> GetContainedDefinitions(IElement container) {
            foreach (IElement definition in _orderedDefinitions) {
                if (IsDescendantOf(definition, container)) yield return definition;
            }
        }

        internal string GetDefinitionFallbackText(IElement definition) {
            var builder = new StringBuilder();
            AppendDefinitionFallbackText(definition, builder);
            return NormalizeBlockText(builder.ToString());
        }

        private static IEnumerable<IElement> EnumerateElements(INode root) {
            foreach (INode child in root.ChildNodes) {
                if (child is IElement element) yield return element;
                foreach (IElement descendant in EnumerateElements(child)) yield return descendant;
            }
        }

        private static bool IsSemanticFootnoteDefinition(IElement element) =>
            HasAnyEpubType(element, "footnote", "endnote", "rearnote", "note")
            || HasAnyRole(element, "doc-footnote", "doc-endnote");

        private static bool IsRecognizedIdDefinition(IElement element, ISet<string> referencedIds) {
            string? id = element.Id;
            if (string.IsNullOrWhiteSpace(id)
                || (!id!.StartsWith("fn:", StringComparison.OrdinalIgnoreCase)
                    && !id.StartsWith("fn-", StringComparison.OrdinalIgnoreCase))) {
                return false;
            }
            if (referencedIds.Contains(id)) return true;
            for (IElement? current = element.ParentElement; current != null; current = current.ParentElement) {
                if (IsFootnoteContainer(current)) return true;
            }
            return false;
        }

        private static bool IsFootnoteContainer(IElement element) =>
            element.HasAttribute("data-footnotes")
            || element.ClassList.Contains("footnotes")
            || element.ClassList.Contains("endnotes")
            || HasAnyEpubType(element, "footnotes", "endnotes", "rearnotes")
            || HasAnyRole(element, "doc-endnotes");

        private static bool LooksLikeFootnoteReference(IElement element) {
            if (element == null || !element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)) return false;
            if (HasAnyEpubType(element, "noteref")
                || HasAnyRole(element, "doc-noteref")
                || element.HasAttribute("data-footnote-ref")
                || element.ClassList.Contains("noteref")
                || element.ClassList.Contains("footnote-ref")) {
                return true;
            }

            IElement? parent = element.ParentElement;
            string? parentId = parent?.Id;
            return parent != null
                   && parent.TagName.Equals("SUP", StringComparison.OrdinalIgnoreCase)
                   && (!string.IsNullOrWhiteSpace(parentId)
                       && (parentId!.StartsWith("fnref:", StringComparison.OrdinalIgnoreCase)
                           || parentId.StartsWith("fnref-", StringComparison.OrdinalIgnoreCase)));
        }

        private static string? GetLocalFragmentId(string? href) {
            if (string.IsNullOrWhiteSpace(href)) return null;
            string candidate = href!.Trim();
            if (!candidate.StartsWith("#", StringComparison.Ordinal) || candidate.Length == 1) return null;
            try {
                return Uri.UnescapeDataString(candidate.Substring(1));
            } catch {
                return candidate.Substring(1);
            }
        }

        private static string CreateUniqueLabel(string? id, int index, ISet<string> usedLabels) {
            string candidate = id?.Trim() ?? string.Empty;
            if (candidate.StartsWith("fn:", StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith("fn-", StringComparison.OrdinalIgnoreCase)) {
                candidate = candidate.Substring(3);
            }
            candidate = SanitizeLabel(candidate);
            if (candidate.Length == 0) candidate = "note-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);

            string unique = candidate;
            int suffix = 2;
            while (!usedLabels.Add(unique)) {
                unique = candidate + "-" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                suffix++;
            }
            return unique;
        }

        private static string SanitizeLabel(string value) {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            var builder = new StringBuilder(value.Length);
            bool previousSeparator = false;
            foreach (char character in value) {
                bool invalid = char.IsWhiteSpace(character)
                    || char.IsControl(character)
                    || character == '['
                    || character == ']'
                    || character == '^';
                if (invalid) {
                    if (!previousSeparator && builder.Length > 0) builder.Append('-');
                    previousSeparator = true;
                    continue;
                }
                builder.Append(character);
                previousSeparator = false;
            }
            return builder.ToString().Trim('-');
        }

        private static string NormalizeComparisonText(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return string.Join(" ", value!.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private void AppendDefinitionFallbackText(INode node, StringBuilder builder) {
            foreach (INode child in node.ChildNodes) {
                if (child is IElement element) {
                    if (IsBacklink(element)) continue;
                    AppendDefinitionFallbackText(element, builder);
                } else if (child is IText text) {
                    builder.Append(text.Data);
                }
            }
        }

        private static bool IsDescendantOf(IElement element, IElement ancestor) {
            for (IElement? current = element.ParentElement; current != null; current = current.ParentElement) {
                if (ReferenceEquals(current, ancestor)) return true;
            }
            return false;
        }

        private static bool HasAnyEpubType(IElement element, params string[] values) =>
            values.Any(value => HtmlAccessibilitySemantics.HasEpubType(element, value));

        private static bool HasAnyRole(IElement element, params string[] values) =>
            values.Any(value => HtmlAccessibilitySemantics.HasRole(element, value));
    }

    private static bool TryConvertFootnoteElement(
        IElement element,
        ConversionContext context,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        blocks = Array.Empty<IMarkdownBlock>();
        if (context.Footnotes.TryGetDefinitionLabel(element, out string label)) {
            IReadOnlyList<IMarkdownBlock> body = ConvertNodesToBlocks(element.ChildNodes, context);
            blocks = new IMarkdownBlock[] {
                body.Count == 0
                    ? new FootnoteDefinitionBlock(label, context.Footnotes.GetDefinitionFallbackText(element))
                    : new FootnoteDefinitionBlock(label, body)
            };
            return true;
        }

        if (!context.Footnotes.ShouldConvertContainer(element)) return false;
        var definitions = new List<IMarkdownBlock>();
        foreach (IElement definition in context.Footnotes.GetContainedDefinitions(element)) {
            if (!context.Footnotes.TryGetDefinitionLabel(definition, out label)) continue;
            IReadOnlyList<IMarkdownBlock> body = ConvertNodesToBlocks(definition.ChildNodes, context);
            definitions.Add(body.Count == 0
                ? new FootnoteDefinitionBlock(label, context.Footnotes.GetDefinitionFallbackText(definition))
                : new FootnoteDefinitionBlock(label, body));
        }
        blocks = definitions;
        return definitions.Count > 0;
    }
}
