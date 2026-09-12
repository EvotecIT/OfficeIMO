using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

public sealed partial class HtmlScriptCapture {
    /// <summary>
    /// Creates an independent frozen HTML document whose first base href is absolute.
    /// Use this snapshot for serialization or conversion to preserve resource resolution
    /// after route changes. The authored <see cref="Document"/> remains unchanged.
    /// </summary>
    public HtmlDocument CreateStandaloneDocument() {
        var clone=Document.CloneAttached();
        var element=FindBase(clone);
        if(element==null) {
            var root=clone.DocumentElement;
            if(root is not {LocalName:"html",NamespaceUri:HtmlElement.HtmlNamespace})
                throw new InvalidOperationException("A standalone capture requires an HTML document root.");
            var head=clone.Head;
            if(head==null) {head=clone.CreateElement("head");Prepend(root,head);}
            element=clone.CreateElement("base");Prepend(head,element);
        }
        element.SetAttribute("href",BaseUri.AbsoluteUri);
        return clone.Freeze();
    }

    private static HtmlElement? FindBase(HtmlDocument document) => document.Descendants().OfType<HtmlElement>()
        .FirstOrDefault(element=>element.NamespaceUri==HtmlElement.HtmlNamespace && element.LocalName=="base" && element.HasAttribute("href"));

    private static Uri ResolveBaseUri(HtmlDocument document,Uri documentUrl,Uri? supplied) {
        if(supplied!=null) {
            if(!supplied.IsAbsoluteUri || supplied.AbsoluteUri.Length>8192)throw new ArgumentException("The captured base URI must be absolute and bounded.",nameof(supplied));
            return supplied;
        }
        return Uri.TryCreate(documentUrl,FindBase(document)?.GetAttribute("href"),out var resolved)
            && resolved.Scheme is not ("data" or "javascript") ? resolved : documentUrl;
    }

    private static void Prepend(HtmlElement parent,HtmlElement child) {
        var previous=parent.ChildNodes.ToArray();
        parent.AppendChild(child);
        foreach(var node in previous)parent.AppendChild(node);
    }
}
