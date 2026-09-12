using System.Runtime.CompilerServices;
using System.Threading;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html;

// Native conversion trees carry the same immutable state as their owned source.
// Cloning must explicitly copy it because provider clones retain authored defaults.
internal static class NativeFormState {
    private static readonly ConditionalWeakTable<IElement, HtmlFormControlState> States = new ConditionalWeakTable<IElement, HtmlFormControlState>();

    internal static HtmlFormControlState? Get(IElement element) => States.TryGetValue(element, out var state) ? state : null;

    internal static void Attach(IElement element, HtmlFormControlState? state) {
        if (state != null) States.Add(element, state);
    }

    internal static void ApplyTree(INode root, CancellationToken cancellationToken = default) {
        var selects = new List<IElement>();
        foreach (INode node in Walk(root, cancellationToken)) {
            if (node is IElement element && Get(element) is HtmlFormControlState state) Apply(element, state);
            if (node is IHtmlSelectElement select) selects.Add(select);
        }
        foreach (IElement select in selects) {
            cancellationToken.ThrowIfCancellationRequested();
            IElement[]? selected = GetSelectedOptions(select);
            if (selected == null) continue;
            var selectedSet = new HashSet<IElement>(selected);
            foreach (IHtmlOptionElement option in select.QuerySelectorAll("option").OfType<IHtmlOptionElement>())
                option.IsSelected = selectedSet.Contains(option);
        }
    }

    // Null means a static select with no live state. An empty result means an explicit empty
    // selection. Current single-select constraints take precedence over retained option flags.
    internal static IElement[]? GetSelectedOptions(IElement select) {
        IElement[] options = select.QuerySelectorAll("option").ToArray();
        if (Get(select) == null && !options.Any(option => Get(option) != null)) return null;
        bool multiple = select.HasAttribute("multiple");
        IElement? staticSelection = null;
        if (Get(select) == null && !multiple) {
            staticSelection = options.LastOrDefault(option => option.HasAttribute("selected"));
            if (staticSelection == null && select is IHtmlSelectElement control && control.Size <= 1)
                staticSelection = options.FirstOrDefault(option => !IsDisabledOption(option));
        }
        IElement[] selected = options.Where(option => Get(option)?.IsSelected
            ?? (Get(select) == null && !multiple ? ReferenceEquals(option, staticSelection) : option.HasAttribute("selected"))).ToArray();
        if (select.HasAttribute("multiple") || selected.Length <= 1) return selected;
        return new[] { selected[selected.Length - 1] };
    }

    private static bool IsDisabledOption(IElement option) {
        if (option.HasAttribute("disabled")) return true;
        for (IElement? parent = option.ParentElement; parent != null && parent is not IHtmlSelectElement; parent = parent.ParentElement)
            if (parent is IHtmlOptionsGroupElement && parent.HasAttribute("disabled")) return true;
        return false;
    }

    internal static long GetValueCharacterCount(INode root, CancellationToken cancellationToken) {
        long characters = 0;
        foreach (INode node in Walk(root, cancellationToken))
            if (node is IElement element) characters += Get(element)?.Value?.Length ?? 0;
        return characters;
    }

    internal static void CopyTree(INode source, INode target, CancellationToken cancellationToken = default) {
        using var sources = Walk(source, cancellationToken).GetEnumerator();
        using var targets = Walk(target, cancellationToken).GetEnumerator();
        while (sources.MoveNext()) {
            if (!targets.MoveNext()) throw new InvalidOperationException("The provider clone changed the document structure.");
            if (sources.Current is IElement original && targets.Current is IElement clone) Attach(clone, Get(original));
        }
        if (targets.MoveNext()) throw new InvalidOperationException("The provider clone changed the document structure.");
        ApplyTree(target, cancellationToken);
    }

    private static void Apply(IElement element, HtmlFormControlState state) {
        if (element is IHtmlInputElement input) {
            if (!string.Equals(input.Value ?? string.Empty, state.Value, StringComparison.Ordinal)) input.Value = state.Value!;
            input.IsChecked = state.IsChecked;
            input.IsIndeterminate = state.IsIndeterminate;
        } else if (element is IHtmlTextAreaElement area) area.Value = state.Value!;
        else if (element is IHtmlOptionElement option) option.IsSelected = state.IsSelected;
    }

    private static IEnumerable<INode> Walk(INode root, CancellationToken cancellationToken) {
        var pending = new Stack<INode>();
        pending.Push(root);
        while (pending.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            INode node = pending.Pop();
            yield return node;
            if (node is IHtmlTemplateElement template) pending.Push(template.Content);
            for (int i = node.ChildNodes.Length - 1; i >= 0; i--) pending.Push(node.ChildNodes[i]);
        }
    }
}
