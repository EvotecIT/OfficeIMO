using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Opml;

/// <summary>An OPML outline element backed by its lossless XML node.</summary>
public sealed class OpmlOutline {
    private readonly OpmlDocument _document;
    internal XElement Element { get; }

    internal OpmlOutline(OpmlDocument document, XElement element) {
        _document = document;
        Element = element;
    }

    /// <summary>Required outline label.</summary>
    public string Text { get => Get("text") ?? string.Empty; set => Set("text", value ?? throw new ArgumentNullException(nameof(value))); }
    /// <summary>Outline type, including subscription types such as rss.</summary>
    public string? Type { get => Get("type"); set => Set("type", value); }
    /// <summary>Subscription feed URL.</summary>
    public string? XmlUrl { get => Get("xmlUrl"); set => Set("xmlUrl", value); }
    /// <summary>Subscription website URL.</summary>
    public string? HtmlUrl { get => Get("htmlUrl"); set => Set("htmlUrl", value); }
    /// <summary>Link/include URL.</summary>
    public string? Url { get => Get("url"); set => Set("url", value); }
    /// <summary>Subscription description.</summary>
    public string? Description { get => Get("description"); set => Set("description", value); }
    /// <summary>Subscription title.</summary>
    public string? Title { get => Get("title"); set => Set("title", value); }
    /// <summary>Subscription language.</summary>
    public string? Language { get => Get("language"); set => Set("language", value); }
    /// <summary>Subscription format version.</summary>
    public string? Version { get => Get("version"); set => Set("version", value); }
    /// <summary>Category path.</summary>
    public string? Category { get => Get("category"); set => Set("category", value); }
    /// <summary>Creation date in producer text form.</summary>
    public string? Created { get => Get("created"); set => Set("created", value); }

    /// <summary>All attributes, including qualified extension attributes.</summary>
    public IReadOnlyDictionary<XName, string> Attributes => Element.Attributes().ToDictionary(a => a.Name, a => a.Value);
    /// <summary>Nested outlines in document order.</summary>
    public IReadOnlyList<OpmlOutline> Children => Element.Elements("outline").Select(e => new OpmlOutline(_document, e)).ToArray();

    /// <summary>Adds a nested outline.</summary>
    public OpmlOutline AddChild(string text) {
        var child = new XElement("outline", new XAttribute("text", text ?? throw new ArgumentNullException(nameof(text))));
        Element.Add(child);
        _document.MarkModified();
        return new OpmlOutline(_document, child);
    }

    /// <summary>Gets any standard or extension attribute.</summary>
    public string? GetAttribute(XName name) => Element.Attribute(name)?.Value;

    /// <summary>Sets or removes any standard or extension attribute.</summary>
    public void SetAttribute(XName name, string? value) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        Element.SetAttributeValue(name, value);
        _document.MarkModified();
    }

    /// <summary>Removes this outline from its parent.</summary>
    public void Remove() {
        Element.Remove();
        _document.MarkModified();
    }

    private string? Get(string name) => Element.Attribute(name)?.Value;
    private void Set(string name, string? value) {
        Element.SetAttributeValue(name, value);
        _document.MarkModified();
    }
}