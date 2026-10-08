using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    // A style computation reads one stable document. Reuse its selector projections
    // within that operation only; a later computation must observe DOM changes.
    internal sealed class AngleSharpSelectorElementCache {
        private readonly Dictionary<IElement, AngleSharpSelectorElement> _elements = new();
        internal AngleSharpSelectorElement Get(IElement element) {
            if (!_elements.TryGetValue(element, out AngleSharpSelectorElement? projection)) {
                projection = new AngleSharpSelectorElement(element, this);
                _elements.Add(element, projection);
            }
            return projection;
        }
    }

    internal sealed class AngleSharpSelectorElement : OfficeIMO.Html.Css.IHtmlCssSelectorElement {
        private readonly IElement _element;
        private readonly AngleSharpSelectorElementCache _cache;
        private IReadOnlyList<OfficeIMO.Html.Css.HtmlCssSelectorAttributeValue>? _attributes;
        internal AngleSharpSelectorElement(IElement element, AngleSharpSelectorElementCache cache) { _element = element; _cache = cache; }
        public object Identity => _element;
        public string LocalName => _element.LocalName ?? _element.TagName ?? string.Empty;
        public string NamespaceUri => _element.NamespaceUri ?? string.Empty;
        public string Id => _element.Id ?? string.Empty;
        public OfficeIMO.Html.Css.IHtmlCssSelectorElement? ParentElement =>
            _element.ParentElement == null ? null : _cache.Get(_element.ParentElement);
        public object? SiblingParentIdentity => _element.Parent;
        public IReadOnlyList<OfficeIMO.Html.Css.IHtmlCssSelectorElement> GetSiblingElements(
            Action? recordEvaluation,
            CancellationToken cancellationToken) {
            var children = new List<OfficeIMO.Html.Css.IHtmlCssSelectorElement>();
            if (_element.Parent == null) return children;
            foreach (INode child in _element.Parent.ChildNodes) {
                cancellationToken.ThrowIfCancellationRequested();
                recordEvaluation?.Invoke();
                if (child is IElement element) children.Add(_cache.Get(element));
            }
            return children;
        }
        public bool IsDocumentElement => ReferenceEquals(_element.Owner?.DocumentElement, _element);
        public bool HasElementOrTextChild(Action? recordEvaluation, CancellationToken cancellationToken) {
            foreach (INode child in _element.ChildNodes) {
                cancellationToken.ThrowIfCancellationRequested();
                recordEvaluation?.Invoke();
                if (child is IElement) return true;
                if (child.NodeType == NodeType.Text && (child.TextContent ?? string.Empty).Length != 0) return true;
            }
            return false;
        }
        public bool HasClass(string name) => _element.ClassList.Contains(name);
        public IReadOnlyList<OfficeIMO.Html.Css.HtmlCssSelectorAttributeValue> Attributes => _attributes ??= _element.Attributes
            .Select(attribute => new OfficeIMO.Html.Css.HtmlCssSelectorAttributeValue(
                attribute.LocalName ?? attribute.Name, attribute.NamespaceUri ?? string.Empty, attribute.Value))
            .ToArray();
    }

}
