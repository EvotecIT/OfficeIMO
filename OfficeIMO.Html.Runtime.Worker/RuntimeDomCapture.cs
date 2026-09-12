using AngleSharp.Dom;
using AngleSharp.Html.Construction;
using AngleSharp.Html.Dom;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeDomCapture {
    internal static HtmlRuntimeWireDocument Capture(IDocument document, HtmlScriptRequest request, CancellationToken token) {
        var mode = ((IConstructableDocument)document).QuirksMode;
        var result = new HtmlRuntimeWireDocument {
            ProviderId = "AngleSharp.Js/" + typeof(AngleSharp.Js.JsScriptingOptions).Assembly.GetName().Version + "; Jint/" + typeof(Jint.Engine).Assembly.GetName().Version,
            Mode = mode == QuirksMode.On ? HtmlDocumentMode.Quirks : mode == QuirksMode.Limited ? HtmlDocumentMode.LimitedQuirks : HtmlDocumentMode.Standards
        };
        var pending = new Queue<(INode Node, int Parent, int Depth, bool Template)>();
        long dataCharacters = 0;
        void Enqueue(INode node, int parent, int depth, bool template = false) {
            if ((long)result.Nodes.Count + pending.Count >= request.MaxNodes) throw new HtmlScriptRuntimeException("Captured node budget exceeded.");
            pending.Enqueue((node, parent, depth, template));
        }
        void Count(string? value) {
            dataCharacters += value?.Length ?? 0;
            if (dataCharacters > request.MaxOutputCharacters) throw new HtmlScriptRuntimeException("Captured data budget exceeded.");
        }
        foreach (INode child in document.ChildNodes) Enqueue(child, 0, 0);
        while (pending.Count != 0) {
            token.ThrowIfCancellationRequested();
            var entry = pending.Dequeue();
            INode source = entry.Node;
            int depth = entry.Depth + (source is IElement ? 1 : 0);
            if (depth > request.MaxDepth) throw new HtmlScriptRuntimeException("Captured depth budget exceeded.");
            var node = new HtmlRuntimeWireNode { Parent = entry.Parent, IsTemplateContent = entry.Template };
            switch (source) {
                case IElement element:
                    if (element is IHtmlInputElement file && string.Equals(file.Type, "file", StringComparison.OrdinalIgnoreCase)
                        && ((file.Files?.Length ?? 0) != 0 || !string.IsNullOrEmpty(file.Value)))
                        throw new HtmlScriptRuntimeException("Capturing selected files is not supported.");
                    node.Kind = HtmlNodeKind.Element;
                    node.Name = element.LocalName;
                    node.NamespaceUri = element.NamespaceUri ?? string.Empty;
                    node.Prefix = element.Prefix;
                    node.FormState = element switch {
                        IHtmlInputElement input => new HtmlFormControlState(HtmlFormControlStateKind.Input, input.Value ?? string.Empty, input.IsChecked, input.IsIndeterminate),
                        IHtmlTextAreaElement area => new HtmlFormControlState(HtmlFormControlStateKind.TextArea, area.Value ?? string.Empty),
                        IHtmlSelectElement => new HtmlFormControlState(HtmlFormControlStateKind.Select),
                        IHtmlOptionElement option => new HtmlFormControlState(HtmlFormControlStateKind.Option, isSelected: option.IsSelected),
                        _ => null
                    };
                    Count(node.FormState?.Value);
                    foreach (IAttr attribute in element.Attributes) {
                        token.ThrowIfCancellationRequested();
                        Count(attribute.Name); Count(attribute.Value); Count(attribute.NamespaceUri);
                        node.Attributes.Add(new HtmlRuntimeWireAttribute { Name = attribute.Name, Value = attribute.Value, NamespaceUri = attribute.NamespaceUri ?? string.Empty });
                    }
                    break;
                case IDocumentType type:
                    node.Kind = HtmlNodeKind.DocumentType; node.Name = type.Name;
                    node.PublicIdentifier = type.PublicIdentifier; node.SystemIdentifier = type.SystemIdentifier;
                    break;
                case IText text: node.Kind = HtmlNodeKind.Text; node.Data = text.Data; break;
                case IComment comment: node.Kind = HtmlNodeKind.Comment; node.Data = comment.Data; break;
                case IDocumentFragment when entry.Template: node.Kind = HtmlNodeKind.DocumentFragment; break;
                default: throw new HtmlScriptRuntimeException("The live document contains an unsupported node kind.");
            }
            Count(node.Name); Count(node.NamespaceUri); Count(node.Prefix); Count(node.Data); Count(node.PublicIdentifier); Count(node.SystemIdentifier);
            result.Nodes.Add(node);
            int parentIndex = result.Nodes.Count;
            foreach (INode child in source.ChildNodes) Enqueue(child, parentIndex, depth);
            if (source is IHtmlTemplateElement template) Enqueue(template.Content, parentIndex, depth, template: true);
        }
        return result;
    }
}
