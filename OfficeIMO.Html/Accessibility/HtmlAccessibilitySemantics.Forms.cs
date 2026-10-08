using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlAccessibilitySemantics {
    private static string ResolveFormLabel(IElement element, ISet<IElement> resolutionPath, HtmlAccessibleNameContext context, int depth) {
        if (!IsLabelable(element) || element.Owner == null) return string.Empty;
        var names = new List<string>();
        foreach (IElement label in context.GetFormLabels(element)) {
            string name = GetAccessibleName(label, true, false, resolutionPath, context, depth + 1);
            if (name.Length > 0) names.Add(name);
            if (names.Count >= MaximumTokens) break;
        }
        return context.NormalizeText(string.Join(" ", names));
    }

    internal sealed partial class HtmlAccessibleNameContext {
        private readonly Dictionary<IDocument, Dictionary<IElement, List<IElement>>> _formLabels = new();

        internal IEnumerable<IElement> GetFormLabels(IElement element) {
            IDocument document = element.Owner!;
            if (!_formLabels.TryGetValue(document, out var associations)) {
                associations = BuildFormLabelIndex(document);
                _formLabels.Add(document, associations);
            }
            return associations.TryGetValue(element, out var labels) ? labels : Array.Empty<IElement>();
        }

        private Dictionary<IElement, List<IElement>> BuildFormLabelIndex(IDocument document) {
            var firstIds = new Dictionary<string, IElement>(StringComparer.Ordinal);
            var labels = new List<IElement>();
            foreach (IElement candidate in EnumerateLabelElements(document, this)) {
                string? id = candidate.Id;
                if (!string.IsNullOrEmpty(id) && !firstIds.ContainsKey(id!)) firstIds.Add(id!, candidate);
                if (candidate.LocalName == "label" && candidate.NamespaceUri == Dom.HtmlElement.HtmlNamespace) labels.Add(candidate);
            }
            var associations = new Dictionary<IElement, List<IElement>>();
            foreach (IElement label in labels) {
                if (!TryConsumeTraversalWork()) break;
                IElement? target = null;
                if (label.HasAttribute("for")) {
                    firstIds.TryGetValue(label.GetAttribute("for")!, out target);
                    if (target != null && !IsLabelable(target)) target = null;
                } else target = EnumerateLabelElements(label, this).FirstOrDefault(IsLabelable);
                if (target == null) continue;
                if (!associations.TryGetValue(target, out var matches)) associations.Add(target, matches = new List<IElement>());
                if (matches.Count < MaximumTokens) matches.Add(label);
            }
            return associations;
        }
    }

    // Walk incrementally so a large document or label cannot allocate an unbounded selector result.
    private static IEnumerable<IElement> EnumerateLabelElements(INode root, HtmlAccessibleNameContext context) {
        INode? current = root.FirstChild;
        while (current != null) {
            if (!context.TryConsumeTraversalWork()) yield break;
            if (current is IElement element) yield return element;
            if (current.FirstChild != null) { current = current.FirstChild; continue; }
            while (current.NextSibling == null && !ReferenceEquals(current.Parent, root)) {
                current = current.Parent;
                if (current == null) yield break;
            }
            current = current.NextSibling;
        }
    }

    private static bool IsLabelable(IElement element) => element.NamespaceUri == Dom.HtmlElement.HtmlNamespace
        && (element.LocalName is "button" or "meter" or "output" or "progress" or "select" or "textarea"
            || element.LocalName == "input" && !string.Equals(element.GetAttribute("type"), "hidden", StringComparison.OrdinalIgnoreCase));
}
