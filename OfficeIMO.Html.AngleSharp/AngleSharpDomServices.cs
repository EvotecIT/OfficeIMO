using AngleSharp;
using AngleSharp.Dom;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Providers;

/// <summary>Selector and HTML serialization services over owned nodes. Performs no network access or script execution.</summary>
public sealed class AngleSharpDomServices : IHtmlDomServices {
    /// <summary>Shared stateless syntax services.</summary>
    public static AngleSharpDomServices Instance { get; } = new AngleSharpDomServices();
    /// <inheritdoc />
    public string Serialize(HtmlNode node, bool childrenOnly = false) {
        if (node == null) throw new ArgumentNullException(nameof(node));
        try {
            INode native = NativeDomBridge.GetNative(node);
            if (childrenOnly && native is AngleSharp.Html.Dom.IHtmlTemplateElement template) native = template.Content;
            return childrenOnly ? string.Concat(native.ChildNodes.Select(child => child.ToHtml())) : native.ToHtml();
        } catch (DomException error) {
            throw new InvalidOperationException("The HTML tree cannot be serialized: " + error.Message);
        }
    }
    /// <inheritdoc />
    public IReadOnlyList<HtmlElement> QuerySelectorAll(HtmlNode scope, string selector) {
        if (scope == null) throw new ArgumentNullException(nameof(scope));
        if (selector == null) throw new ArgumentNullException(nameof(selector));
        var state = NativeDomBridge.GetState(scope.Document);
        INode native = state.ToNative[scope.NodeId];
        try {
            IEnumerable<IElement> elements = native is IParentNode parent ? parent.QuerySelectorAll(selector) : Array.Empty<IElement>();
            return elements.Select(element => (HtmlElement)state.ToOwned[element]).ToArray();
        } catch (DomException error) {
            throw new ArgumentException("The CSS selector is invalid: " + error.Message, nameof(selector));
        }
    }
    /// <inheritdoc />
    public bool Matches(HtmlElement element, string selector) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (selector == null) throw new ArgumentNullException(nameof(selector));
        try { return ((IElement)NativeDomBridge.GetNative(element)).Matches(selector); }
        catch (DomException error) { throw new ArgumentException("The CSS selector is invalid: " + error.Message, nameof(selector)); }
    }
}
