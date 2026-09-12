using System;

namespace OfficeIMO.Html.Dom;

/// <summary>An immutable attribute, including its namespace and qualified name.</summary>
public sealed class HtmlAttribute {
    /// <summary>Creates an attribute without interpreting its value as markup or a URL.</summary>
    public HtmlAttribute(string name, string value, string? namespaceUri = null) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("An attribute name is required.", nameof(name));
        Name = name;
        Value = value ?? throw new ArgumentNullException(nameof(value));
        NamespaceUri = namespaceUri ?? string.Empty;
        int colon = NamespaceUri.Length == 0 ? -1 : name.IndexOf(':');
        Prefix = colon < 0 ? string.Empty : name.Substring(0, colon);
        LocalName = colon < 0 ? name : name.Substring(colon + 1);
    }

    /// <summary>Qualified attribute name.</summary>
    public string Name { get; }
    /// <summary>Local name without the namespace prefix.</summary>
    public string LocalName { get; }
    /// <summary>Namespace prefix, or an empty string.</summary>
    public string Prefix { get; }
    /// <summary>Namespace URI, or an empty string.</summary>
    public string NamespaceUri { get; }
    /// <summary>Decoded attribute value.</summary>
    public string Value { get; }
}
