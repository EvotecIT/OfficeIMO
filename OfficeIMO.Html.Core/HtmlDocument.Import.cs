using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Html.Dom;

public sealed partial class HtmlDocument {
    /// <summary>
    /// Copies a node into this mutable document and returns the detached copy. The source remains
    /// unchanged. Copies have fresh destination-local IDs and no source offsets; namespaces,
    /// attributes, decoded data and, for deep copies, template contents are preserved.
    /// </summary>
    /// <param name="source">An element, text, comment, document type or fragment from any owned document.</param>
    /// <param name="deep">Whether to include descendants and template contents.</param>
    /// <param name="cancellationToken">Cooperative cancellation while copying nodes.</param>
    /// <remarks>
    /// Import does not insert or sanitize content. Append the returned node explicitly. Importing a
    /// fragment and appending it splices its children. A cancelled import can leave detached partial
    /// copies in this document; use an Edit session when the operation must produce an atomic snapshot.
    /// </remarks>
    public HtmlNode ImportNode(HtmlNode source, bool deep = true, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        EnsureMutable();
        if (source.Kind == HtmlNodeKind.Document) throw new ArgumentException("Import a document's children instead of the document itself.", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlNode result = ImportShallowNode(source, cancellationToken);
        if (!deep) return result;

        var pending = new Stack<(HtmlNode Source, HtmlNode Target)>();
        pending.Push((source, result));
        while (pending.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var pair = pending.Pop();
            foreach (HtmlNode child in pair.Source.ChildNodes) {
                cancellationToken.ThrowIfCancellationRequested();
                HtmlNode copy = ImportShallowNode(child, cancellationToken);
                pair.Target.AppendChild(copy);
                pending.Push((child, copy));
            }
            if (pair.Source is HtmlElement element && element.TemplateContent != null && pair.Target is HtmlElement target) {
                pending.Push((element.TemplateContent, GetOrCreateTemplateContent(target)));
            }
        }
        return result;
    }

    private HtmlNode ImportShallowNode(HtmlNode source, CancellationToken cancellationToken) {
        if (source is HtmlElement element) {
            HtmlElement copy = CreateElement(element.LocalName, element.NamespaceUri, element.Prefix);
            foreach (HtmlAttribute attribute in element.Attributes) {
                cancellationToken.ThrowIfCancellationRequested();
                copy.SetAttribute(attribute.Name, attribute.Value, attribute.NamespaceUri);
            }
            if (element.TemplateContent != null) GetOrCreateTemplateContent(copy);
            return copy;
        }
        if (source is HtmlDocumentType type) return CreateDocumentType(type.Name, type.PublicIdentifier, type.SystemIdentifier);
        if (source.Kind == HtmlNodeKind.Text) return CreateTextNode(source.TextContent);
        if (source.Kind == HtmlNodeKind.Comment) return CreateComment(source.TextContent);
        if (source.Kind == HtmlNodeKind.DocumentFragment) return CreateFragment();
        throw new ArgumentException("The node kind cannot be imported.", nameof(source));
    }
}
