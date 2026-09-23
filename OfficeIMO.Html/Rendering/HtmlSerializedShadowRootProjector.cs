using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Composes Chromium's MHTML shadow-root serialization on the private render clone.
/// The resulting light DOM is intentionally a static approximation of slot distribution.
/// </summary>
internal static class HtmlSerializedShadowRootProjector {
    internal static void Apply(IHtmlDocument document, HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) {
        if (options.ProjectSerializedShadowRoots != true || document.DocumentElement == null) return;

        int projected = 0;
        int omittedStylesheets = 0;
        int visited = 0;
        var pending = new Stack<IElement>();
        pending.Push(document.DocumentElement);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            IElement host = pending.Pop();
            if (++visited > options.MaxHtmlNodes) {
                throw new HtmlDomLimitException(HtmlRenderDiagnosticCodes.NodeLimitExceeded,
                    "Serialized shadow-root projection exceeded the configured DOM node limit.",
                    nameof(HtmlRenderOptions.MaxHtmlNodes), visited, options.MaxHtmlNodes);
            }

            IHtmlTemplateElement? snapshot = host.Children
                .OfType<IHtmlTemplateElement>()
                .FirstOrDefault(template => IsShadowMode(template.GetAttribute("shadowmode")));
            if (snapshot != null) {
                omittedStylesheets += ComposeHost(host, snapshot, cancellationToken);
                projected++;
            }

            for (int index = host.Children.Length - 1; index >= 0; index--) {
                IElement child = host.Children[index];
                if (child is not IHtmlTemplateElement) pending.Push(child);
            }
        }

        if (projected == 0) return;
        diagnostics.Add("OfficeIMO.Html", HtmlRenderDiagnosticCodes.SerializedShadowRootApproximated,
            "Serialized shadow-root content was projected for static rendering; scoped styling and live behavior may differ.",
            detail: "hosts=" + projected, lossKind: OfficeConversionLossKind.Approximation);
        if (omittedStylesheets > 0) {
            diagnostics.Add("OfficeIMO.Html", HtmlRenderDiagnosticCodes.SerializedShadowStyleOmitted,
                "Shadow-scoped stylesheets were omitted rather than applied to unrelated document content.",
                detail: "stylesheets=" + omittedStylesheets, lossKind: OfficeConversionLossKind.Omission);
        }
    }

    private static bool IsShadowMode(string? value) =>
        string.Equals(value, "open", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "closed", StringComparison.OrdinalIgnoreCase);

    private static int ComposeHost(IElement host, IHtmlTemplateElement snapshot,
        CancellationToken cancellationToken) {
        INode[] lightChildren = host.ChildNodes.Where(child => !ReferenceEquals(child, snapshot)).ToArray();
        INode[] shadowChildren = snapshot.Content.ChildNodes.ToArray();
        foreach (INode child in host.ChildNodes.ToArray()) host.RemoveChild(child);
        foreach (INode child in shadowChildren) host.AppendChild(child);
        int omittedStylesheets = RemoveScopedStyles(shadowChildren, cancellationToken);

        IElement[] slots = host.QuerySelectorAll("slot").ToArray();
        var assigned = new HashSet<INode>();
        var filledNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (IElement slot in slots) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsDescendantOf(slot, host) || slot.Parent == null) continue;
            string name = slot.GetAttribute("name") ?? string.Empty;
            INode[] content = !filledNames.Add(name)
                ? Array.Empty<INode>()
                : lightChildren.Where(child => !assigned.Contains(child)
                    && IsSlotAssignable(child)
                    && string.Equals(SlotName(child), name, StringComparison.Ordinal)).ToArray();
            if (content.Length > 0) {
                foreach (INode child in content) assigned.Add(child);
            } else {
                content = slot.ChildNodes.ToArray();
            }
            INode parent = slot.Parent;
            foreach (INode child in content) parent.InsertBefore(child, slot);
            parent.RemoveChild(slot);
        }
        return omittedStylesheets;
    }

    private static int RemoveScopedStyles(IEnumerable<INode> shadowChildren,
        CancellationToken cancellationToken) {
        int omitted = 0;
        var pending = new Stack<INode>(shadowChildren.Reverse());
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            INode node = pending.Pop();
            if (node is not IElement element) continue;
            if (element.TagName.Equals("style", StringComparison.OrdinalIgnoreCase)
                || element.TagName.Equals("link", StringComparison.OrdinalIgnoreCase)
                && (element.GetAttribute("rel") ?? string.Empty).Split((char[]?)null,
                    StringSplitOptions.RemoveEmptyEntries).Contains("stylesheet", StringComparer.OrdinalIgnoreCase)) {
                node.Parent?.RemoveChild(node);
                omitted++;
                continue;
            }
            for (int index = element.ChildNodes.Length - 1; index >= 0; index--)
                pending.Push(element.ChildNodes[index]);
        }
        return omitted;
    }

    private static bool IsSlotAssignable(INode node) => node is IElement || node is IText;

    private static string SlotName(INode node) => node is IElement element
        ? element.GetAttribute("slot") ?? string.Empty
        : string.Empty;

    private static bool IsDescendantOf(INode node, INode ancestor) {
        for (INode? current = node.Parent; current != null; current = current.Parent) {
            if (ReferenceEquals(current, ancestor)) return true;
        }
        return false;
    }
}
