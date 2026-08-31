using System;
using System.Linq;
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
        if (value == null) {
            child?.Remove();
            _document.MarkModified();
            return;
        }
        if (child == null) child = new XElement(name);
        child.Value = value;
        InsertElement(child);
    }

    internal void InsertElement(XElement element) {
        if (element.Parent != null) element.Remove();
        int order = GetStandardElementOrder(element.Name);
        XElement? following = order < 0 ? null : _element.Elements().FirstOrDefault(candidate =>
            GetStandardElementOrder(candidate.Name) > order);
        if (following == null) _element.Add(element); else following.AddBeforeSelf(element);
        _document.MarkModified();
    }

    internal static int GetStandardElementOrder(XName name) {
        if (name.Namespace != XNamespace.None) return -1;
        switch (name.LocalName) {
            case "title": return 0;
            case "dateCreated": return 1;
            case "dateModified": return 2;
            case "ownerName": return 3;
            case "ownerEmail": return 4;
            case "ownerId": return 5;
            case "docs": return 6;
            case "expansionState": return 7;
            case "vertScrollState": return 8;
            case "windowTop": return 9;
            case "windowLeft": return 10;
            case "windowBottom": return 11;
            case "windowRight": return 12;
            default: return -1;
        }
    }
}
