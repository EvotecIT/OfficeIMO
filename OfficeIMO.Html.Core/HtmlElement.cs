using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Html.Dom;

/// <summary>An owned element with ordered, namespace-aware attributes.</summary>
public sealed class HtmlElement : HtmlNode {
    private readonly List<HtmlAttribute> _attributes = new List<HtmlAttribute>();
    private readonly ReadOnlyCollection<HtmlAttribute> _attributesView;
    private HtmlFormControlState? _formState;

    internal HtmlElement(HtmlDocument document, int id, string name, string namespaceUri, string? prefix = null) : base(document, id, HtmlNodeKind.Element) {
        LocalName = name;
        NamespaceUri = namespaceUri;
        Prefix = prefix ?? string.Empty;
        _attributesView = _attributes.AsReadOnly();
    }
    /// <summary>HTML namespace URI.</summary>
    public const string HtmlNamespace = "http://www.w3.org/1999/xhtml";
    /// <summary>Element local name, preserving foreign-content case.</summary>
    public string LocalName { get; }
    /// <summary>Element namespace URI.</summary>
    public string NamespaceUri { get; }
    /// <summary>Namespace prefix, or an empty string.</summary>
    public string Prefix { get; }
    /// <inheritdoc />
    public override string NodeName => Prefix.Length == 0 ? LocalName : Prefix + ":" + LocalName;
    /// <summary>Ordered attributes.</summary>
    public IReadOnlyList<HtmlAttribute> Attributes => _attributesView;
    /// <summary>Value of the empty-namespace id attribute, or an empty string.</summary>
    public string Id => GetAttribute(string.Empty, "id") ?? string.Empty;
    /// <summary>Value of the empty-namespace class attribute, or an empty string.</summary>
    public string ClassName => GetAttribute(string.Empty, "class") ?? string.Empty;
    /// <summary>Class tokens for inspection.</summary>
    public IReadOnlyList<string> ClassList => ClassName.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries);
    /// <summary>Template contents, retained separately from normal child nodes.</summary>
    public HtmlNode? TemplateContent { get; internal set; }
    /// <summary>Optional immutable live form properties. Attributes and text retain their authored defaults.</summary>
    /// <remarks>Assignment requires a mutable document and a matching HTML form element.
    /// Set to null to use attribute-based form semantics again.</remarks>
    public HtmlFormControlState? FormState {
        get => _formState;
        set {
            Document.EnsureMutable();
            if (value != null && (NamespaceUri != HtmlNamespace || LocalName != value.ElementName))
                throw new ArgumentException("The form state must match this HTML element.", nameof(value));
            _formState = value;
            Document.Touch();
        }
    }

    /// <summary>Reads the first attribute with this qualified name, regardless of namespace.</summary>
    public string? GetAttribute(string name) => _attributes.FirstOrDefault(attribute => NamesEqual(attribute.Name, name))?.Value;
    /// <summary>Checks whether a qualified attribute is present.</summary>
    public bool HasAttribute(string name) => _attributes.Any(attribute => NamesEqual(attribute.Name, name));
    /// <summary>Reads an attribute by namespace URI and local name.</summary>
    public string? GetAttribute(string namespaceUri, string localName) => _attributes.FirstOrDefault(attribute => attribute.NamespaceUri == namespaceUri && attribute.LocalName == localName)?.Value;
    /// <summary>Adds or replaces an attribute in the supplied namespace; omitting the namespace selects the empty namespace.</summary>
    public void SetAttribute(string name, string value, string? namespaceUri = null) {
        Document.EnsureMutable();
        if (NamespaceUri == HtmlNamespace && string.IsNullOrEmpty(namespaceUri)) name = HtmlNames.LowerAscii(name);
        SetAttribute(new HtmlAttribute(name, value, namespaceUri));
    }
    /// <summary>Adds or replaces an exact attribute by namespace and local name, preserving its qualified name and case.</summary>
    /// <remarks>Use this overload when importing structural DOM state, including namespace-aware script mutations.</remarks>
    public void SetAttribute(HtmlAttribute replacement) {
        Document.EnsureMutable();
        if (replacement == null) throw new ArgumentNullException(nameof(replacement));
        int index = _attributes.FindIndex(attribute => attribute.NamespaceUri == replacement.NamespaceUri && attribute.LocalName == replacement.LocalName);
        if (index >= 0) _attributes[index] = replacement; else _attributes.Add(replacement);
        Document.Touch();
    }
    /// <summary>Removes the first attribute with this qualified name, regardless of namespace.</summary>
    public void RemoveAttribute(string name) {
        Document.EnsureMutable();
        int index = _attributes.FindIndex(attribute => NamesEqual(attribute.Name, name));
        if (index >= 0) { _attributes.RemoveAt(index); Document.Touch(); }
    }
    /// <summary>Removes an attribute by namespace URI and local name.</summary>
    public void RemoveAttribute(string namespaceUri, string localName) {
        Document.EnsureMutable();
        if (_attributes.RemoveAll(attribute => attribute.NamespaceUri == namespaceUri && attribute.LocalName == localName) != 0) Document.Touch();
    }
    /// <summary>Tests this element against a CSS selector.</summary>
    public bool Matches(string selector) => Document.Services.Matches(this, selector ?? throw new ArgumentNullException(nameof(selector)));
    /// <summary>Finds this element or its nearest ancestor matching a selector.</summary>
    public HtmlElement? Closest(string selector) {
        for (HtmlElement? current = this; current != null; current = current.ParentElement) if (current.Matches(selector)) return current;
        return null;
    }
    private bool NamesEqual(string left, string right) => string.Equals(left, NamespaceUri == HtmlNamespace ? HtmlNames.LowerAscii(right) : right, StringComparison.Ordinal);
}
