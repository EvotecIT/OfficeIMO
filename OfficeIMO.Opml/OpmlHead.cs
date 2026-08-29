using System;
using System.Xml.Linq;

namespace OfficeIMO.Opml;

/// <summary>Typed access to standard OPML head values while retaining unknown elements.</summary>
public sealed class OpmlHead {
    private readonly OpmlDocument _document;
    private readonly XElement _element;

    internal OpmlHead(OpmlDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>Document title.</summary>
    public string? Title { get => Get("title"); set => Set("title", value); }
    /// <summary>Creation date in the OPML producer's original textual form.</summary>
    public string? DateCreated { get => Get("dateCreated"); set => Set("dateCreated", value); }
    /// <summary>Modification date in the OPML producer's original textual form.</summary>
    public string? DateModified { get => Get("dateModified"); set => Set("dateModified", value); }
    /// <summary>Owner name.</summary>
    public string? OwnerName { get => Get("ownerName"); set => Set("ownerName", value); }
    /// <summary>Owner email address.</summary>
    public string? OwnerEmail { get => Get("ownerEmail"); set => Set("ownerEmail", value); }
    /// <summary>Documentation URL.</summary>
    public string? Docs { get => Get("docs"); set => Set("docs", value); }
    /// <summary>Comma-separated expansion-state list.</summary>
    public string? ExpansionState { get => Get("expansionState"); set => Set("expansionState", value); }

    private string? Get(string name) => _element.Element(name)?.Value;

    private void Set(string name, string? value) {
        XElement? child = _element.Element(name);
        if (value == null) child?.Remove();
        else if (child == null) _element.Add(new XElement(name, value));
        else child.Value = value;
        _document.MarkModified();
    }
}