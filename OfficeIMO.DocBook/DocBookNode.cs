using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.DocBook;

/// <summary>
/// Typed common-structure view over a DocBook element. Unknown elements and attributes remain in the backing XML.
/// </summary>
public sealed class DocBookNode {
    private readonly DocBookDocument _document;
    internal XElement Element { get; }

    internal DocBookNode(DocBookDocument document, XElement element) { _document = document; Element = element; }

    /// <summary>Common structure kind, or Unknown for a preserved extension element.</summary>
    public DocBookNodeKind Kind => DocBookNames.GetKind(Element.Name, _document.Namespace);
    /// <summary>Element local name.</summary>
    public string Name => Element.Name.LocalName;
    /// <summary>Combined descendant text. Setting it intentionally replaces child content.</summary>
    public string Text {
        get => Element.Value;
        set {
            if (Kind == DocBookNodeKind.Author) {
                Element.ReplaceNodes(new XElement(_document.Namespace + "personname", value ?? string.Empty));
            } else {
                Element.Value = value ?? string.Empty;
            }
            _document.MarkModified();
        }
    }
    /// <summary>All attributes, including namespaced extension attributes.</summary>
    public IReadOnlyDictionary<XName, string> Attributes => Element.Attributes().ToDictionary(a => a.Name, a => a.Value);
    /// <summary>Child elements in source order.</summary>
    public IReadOnlyList<DocBookNode> Children => Element.Elements().Select(e => new DocBookNode(_document, e)).ToArray();

    /// <summary>Gets any standard or extension attribute.</summary>
    public string? GetAttribute(XName name) => Element.Attribute(name)?.Value;
    /// <summary>Sets or removes any standard or extension attribute.</summary>
    public void SetAttribute(XName name, string? value) { Element.SetAttributeValue(name, value); _document.MarkModified(); }

    /// <summary>Adds a section with a title.</summary>
    public DocBookNode AddSection(string title) {
        DocBookNode section = Add(DocBookNodeKind.Section);
        section.Add(DocBookNodeKind.Title, title);
        return section;
    }
    /// <summary>Adds a paragraph.</summary>
    public DocBookNode AddParagraph(string text) => Add(DocBookNodeKind.Paragraph, text);
    /// <summary>Adds an itemized list.</summary>
    public DocBookNode AddItemizedList() => Add(DocBookNodeKind.ItemizedList);
    /// <summary>Adds an ordered list.</summary>
    public DocBookNode AddOrderedList() => Add(DocBookNodeKind.OrderedList);
    /// <summary>Adds a list item containing a paragraph.</summary>
    public DocBookNode AddListItem(string text) {
        DocBookNode item = Add(DocBookNodeKind.ListItem);
        item.AddParagraph(text);
        return item;
    }
    /// <summary>Adds a table.</summary>
    public DocBookNode AddTable(string? title = null) {
        DocBookNode table = Add(DocBookNodeKind.Table);
        if (!string.IsNullOrEmpty(title)) table.Add(DocBookNodeKind.Title, title);
        return table;
    }
    /// <summary>Adds a program listing.</summary>
    public DocBookNode AddProgramListing(string code, string? language = null) {
        DocBookNode listing = Add(DocBookNodeKind.ProgramListing, code);
        if (!string.IsNullOrEmpty(language)) listing.SetAttribute("language", language);
        return listing;
    }
    /// <summary>Adds an external link.</summary>
    public DocBookNode AddLink(string text, string href) {
        DocBookNode link = _document.Profile == DocBookProfile.DocBook45 ? AddRaw("ulink", text) : Add(DocBookNodeKind.Link, text);
        XName attribute = _document.Profile == DocBookProfile.DocBook52
            ? XName.Get("href", "http://www.w3.org/1999/xlink") : XName.Get("url");
        link.SetAttribute(attribute, href);
        return link;
    }
    /// <summary>Adds an admonition such as note, tip, important, caution, or warning.</summary>
    public DocBookNode AddAdmonition(DocBookNodeKind kind, string text) {
        if (kind != DocBookNodeKind.Note && kind != DocBookNodeKind.Tip && kind != DocBookNodeKind.Important &&
            kind != DocBookNodeKind.Caution && kind != DocBookNodeKind.Warning) throw new ArgumentOutOfRangeException(nameof(kind));
        DocBookNode node = Add(kind);
        node.AddParagraph(text);
        return node;
    }
    /// <summary>Adds an image media object.</summary>
    public DocBookNode AddImage(string fileReference, string? caption = null) => AddImage(fileReference, caption, null);
    /// <summary>Adds an image media object with distinct caption and alternative text.</summary>
    public DocBookNode AddImage(string fileReference, string? caption, string? alternateText) {
        DocBookNode media = Add(DocBookNodeKind.MediaObject);
        DocBookNode imageObject = media.Add(DocBookNodeKind.ImageObject);
        DocBookNode data = imageObject.Add(DocBookNodeKind.ImageData);
        data.SetAttribute("fileref", fileReference);
        if (!string.IsNullOrEmpty(alternateText)) media.AddRaw("textobject").AddRaw("phrase", alternateText!);
        if (!string.IsNullOrEmpty(caption)) media.Add(DocBookNodeKind.Caption).AddParagraph(caption!);
        return media;
    }
    /// <summary>Adds an index term.</summary>
    public DocBookNode AddIndexTerm(string primary) {
        DocBookNode term = Add(DocBookNodeKind.IndexTerm);
        term.AddRaw("primary", primary);
        return term;
    }

    /// <summary>Adds a supported common node.</summary>
    public DocBookNode Add(DocBookNodeKind kind, string? text = null) {
        if (kind == DocBookNodeKind.Unknown) throw new ArgumentOutOfRangeException(nameof(kind));
        if (kind == DocBookNodeKind.Info && Kind == DocBookNodeKind.Section) {
            string infoName = _document.Profile == DocBookProfile.DocBook45 ? "sectioninfo" : "info";
            var element = new XElement(_document.Namespace + infoName);
            if (text != null) element.Value = text;
            Element.AddFirst(element); _document.MarkModified();
            return new DocBookNode(_document, element);
        }
        if (kind == DocBookNodeKind.Info && _document.Profile == DocBookProfile.DocBook45) {
            return AddRaw(Kind == DocBookNodeKind.Section
                ? "sectioninfo"
                : _document.Kind == DocBookDocumentKind.Article ? "articleinfo" : "bookinfo", text);
        }
        if (kind == DocBookNodeKind.Author && text != null) {
            DocBookNode author = AddRaw(DocBookNames.GetElementName(kind));
            author.AddRaw("personname", text);
            return author;
        }
        return AddRaw(DocBookNames.GetElementName(kind), text);
    }

    /// <summary>Adds an extension element without claiming typed vocabulary support.</summary>
    public DocBookNode AddExtension(XName name, string? text = null) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        var element = new XElement(name);
        if (text != null) element.Value = text;
        Element.Add(element); _document.MarkModified();
        return new DocBookNode(_document, element);
    }

    /// <summary>Removes this node.</summary>
    public void Remove() { Element.Remove(); _document.MarkModified(); }

    internal DocBookNode AddRaw(string localName, string? text = null) {
        XName name = _document.Namespace + localName;
        var element = new XElement(name);
        if (text != null) element.Value = text;
        _document.ResolveTypedContentParent(Element, localName).Add(element); _document.MarkModified();
        return new DocBookNode(_document, element);
    }

    internal void AddText(string text) {
        if (Kind == DocBookNodeKind.Author) {
            XElement? personName = Element.Element(_document.Namespace + "personname");
            if (personName == null) {
                personName = new XElement(_document.Namespace + "personname");
                Element.AddFirst(personName);
            }
            personName.Add(new XText(text ?? string.Empty));
        } else {
            Element.Add(new XText(text ?? string.Empty));
        }
        _document.MarkModified();
    }
}

internal static class DocBookNames {
    private static readonly IReadOnlyDictionary<string, DocBookNodeKind> Kinds = new Dictionary<string, DocBookNodeKind>(StringComparer.Ordinal) {
        ["info"] = DocBookNodeKind.Info,
        ["articleinfo"] = DocBookNodeKind.Info,
        ["bookinfo"] = DocBookNodeKind.Info,
        ["sectioninfo"] = DocBookNodeKind.Info,
        ["title"] = DocBookNodeKind.Title,
        ["subtitle"] = DocBookNodeKind.Subtitle,
        ["author"] = DocBookNodeKind.Author,
        ["section"] = DocBookNodeKind.Section,
        ["sect1"] = DocBookNodeKind.Section,
        ["sect2"] = DocBookNodeKind.Section,
        ["sect3"] = DocBookNodeKind.Section,
        ["sect4"] = DocBookNodeKind.Section,
        ["sect5"] = DocBookNodeKind.Section,
        ["para"] = DocBookNodeKind.Paragraph,
        ["simpara"] = DocBookNodeKind.Paragraph,
        ["itemizedlist"] = DocBookNodeKind.ItemizedList,
        ["orderedlist"] = DocBookNodeKind.OrderedList,
        ["variablelist"] = DocBookNodeKind.VariableList,
        ["listitem"] = DocBookNodeKind.ListItem,
        ["table"] = DocBookNodeKind.Table,
        ["informaltable"] = DocBookNodeKind.Table,
        ["tgroup"] = DocBookNodeKind.TableGroup,
        ["thead"] = DocBookNodeKind.TableHead,
        ["tbody"] = DocBookNodeKind.TableBody,
        ["row"] = DocBookNodeKind.Row,
        ["entry"] = DocBookNodeKind.Entry,
        ["programlisting"] = DocBookNodeKind.ProgramListing,
        ["screen"] = DocBookNodeKind.Screen,
        ["link"] = DocBookNodeKind.Link,
        ["ulink"] = DocBookNodeKind.Link,
        ["xref"] = DocBookNodeKind.CrossReference,
        ["note"] = DocBookNodeKind.Note,
        ["tip"] = DocBookNodeKind.Tip,
        ["important"] = DocBookNodeKind.Important,
        ["caution"] = DocBookNodeKind.Caution,
        ["warning"] = DocBookNodeKind.Warning,
        ["figure"] = DocBookNodeKind.Figure,
        ["mediaobject"] = DocBookNodeKind.MediaObject,
        ["imageobject"] = DocBookNodeKind.ImageObject,
        ["imagedata"] = DocBookNodeKind.ImageData,
        ["caption"] = DocBookNodeKind.Caption,
        ["index"] = DocBookNodeKind.Index,
        ["indexterm"] = DocBookNodeKind.IndexTerm
    };

    internal static DocBookNodeKind GetKind(string name) => Kinds.TryGetValue(name, out DocBookNodeKind kind) ? kind : DocBookNodeKind.Unknown;
    internal static DocBookNodeKind GetKind(XName name, XNamespace docBookNamespace) =>
        name.Namespace == docBookNamespace ? GetKind(name.LocalName) : DocBookNodeKind.Unknown;
    internal static string GetElementName(DocBookNodeKind kind) {
        switch (kind) {
            case DocBookNodeKind.Info: return "info";
            case DocBookNodeKind.Title: return "title";
            case DocBookNodeKind.Subtitle: return "subtitle";
            case DocBookNodeKind.Author: return "author";
            case DocBookNodeKind.Section: return "section";
            case DocBookNodeKind.Paragraph: return "para";
            case DocBookNodeKind.ItemizedList: return "itemizedlist";
            case DocBookNodeKind.OrderedList: return "orderedlist";
            case DocBookNodeKind.VariableList: return "variablelist";
            case DocBookNodeKind.ListItem: return "listitem";
            case DocBookNodeKind.Table: return "table";
            case DocBookNodeKind.TableGroup: return "tgroup";
            case DocBookNodeKind.TableHead: return "thead";
            case DocBookNodeKind.TableBody: return "tbody";
            case DocBookNodeKind.Row: return "row";
            case DocBookNodeKind.Entry: return "entry";
            case DocBookNodeKind.ProgramListing: return "programlisting";
            case DocBookNodeKind.Screen: return "screen";
            case DocBookNodeKind.Link: return "link";
            case DocBookNodeKind.CrossReference: return "xref";
            case DocBookNodeKind.Note: return "note";
            case DocBookNodeKind.Tip: return "tip";
            case DocBookNodeKind.Important: return "important";
            case DocBookNodeKind.Caution: return "caution";
            case DocBookNodeKind.Warning: return "warning";
            case DocBookNodeKind.Figure: return "figure";
            case DocBookNodeKind.MediaObject: return "mediaobject";
            case DocBookNodeKind.ImageObject: return "imageobject";
            case DocBookNodeKind.ImageData: return "imagedata";
            case DocBookNodeKind.Caption: return "caption";
            case DocBookNodeKind.Index: return "index";
            case DocBookNodeKind.IndexTerm: return "indexterm";
            default: throw new ArgumentOutOfRangeException(nameof(kind));
        }
    }
}
