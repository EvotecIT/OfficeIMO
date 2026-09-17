using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Html.Construction;
using AngleSharp.Html.Dom;
using OfficeIMO.Html.Dom;
using System.Reflection;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeDomCapture {
    private const int MaxFrameDepth = 8;

    internal static HtmlRuntimeWireDocument Capture(IDocument document, HtmlScriptRequest request, CancellationToken token) {
        var budget = new CaptureBudget(request);
        return CaptureDocument(document, request, token, budget, frameDepth: 0, inheritedDocumentUrl: null);
    }

    private static HtmlRuntimeWireDocument CaptureDocument(
        IDocument document,
        HtmlScriptRequest request,
        CancellationToken token,
        CaptureBudget budget,
        int frameDepth,
        Uri? inheritedDocumentUrl) {
        var mode = ((IConstructableDocument)document).QuirksMode;
        Uri documentUrl = ResolveDocumentUrl(document, inheritedDocumentUrl);
        var result = new HtmlRuntimeWireDocument {
            ProviderId = "AngleSharp/" + typeof(IDocument).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion
                + "; AngleSharp.Js/" + typeof(AngleSharp.Js.JsScriptingOptions).Assembly.GetName().Version + "; Jint/" + typeof(Jint.Engine).Assembly.GetName().Version,
            Mode = mode == QuirksMode.On ? HtmlDocumentMode.Quirks : mode == QuirksMode.Limited ? HtmlDocumentMode.LimitedQuirks : HtmlDocumentMode.Standards,
            DocumentUrl = documentUrl,
            BaseUri = ResolveBaseUri(document, documentUrl)
        };
        var pending = new Queue<(INode Node, int Parent, int Depth, bool Template)>();
        var frameNodeIds = new Dictionary<IHtmlInlineFrameElement, int>();

        void Enqueue(INode node, int parent, int depth, bool template = false) {
            budget.ReserveNode();
            pending.Enqueue((node, parent, depth, template));
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
                    budget.Count(node.FormState?.Value);
                    foreach (IAttr attribute in element.Attributes) {
                        token.ThrowIfCancellationRequested();
                        budget.Count(attribute.Name);
                        budget.Count(attribute.Value);
                        budget.Count(attribute.NamespaceUri);
                        node.Attributes.Add(new HtmlRuntimeWireAttribute {
                            Name = attribute.Name,
                            Value = attribute.Value ?? string.Empty,
                            NamespaceUri = attribute.NamespaceUri ?? string.Empty
                        });
                    }
                    break;
                case IDocumentType type:
                    node.Kind = HtmlNodeKind.DocumentType;
                    node.Name = type.Name;
                    node.PublicIdentifier = type.PublicIdentifier;
                    node.SystemIdentifier = type.SystemIdentifier;
                    break;
                case IText text:
                    node.Kind = HtmlNodeKind.Text;
                    node.Data = text.Data;
                    break;
                case IComment comment:
                    node.Kind = HtmlNodeKind.Comment;
                    node.Data = comment.Data;
                    break;
                case IDocumentFragment when entry.Template:
                    node.Kind = HtmlNodeKind.DocumentFragment;
                    break;
                default:
                    throw new HtmlScriptRuntimeException("The live document contains an unsupported node kind.");
            }
            budget.Count(node.Name);
            budget.Count(node.NamespaceUri);
            budget.Count(node.Prefix);
            budget.Count(node.Data);
            budget.Count(node.PublicIdentifier);
            budget.Count(node.SystemIdentifier);
            result.Nodes.Add(node);
            int nodeId = result.Nodes.Count;
            if (source is IHtmlInlineFrameElement frame) frameNodeIds.Add(frame, nodeId);
            foreach (INode child in source.ChildNodes) Enqueue(child, nodeId, depth);
            if (source is IHtmlTemplateElement template) Enqueue(template.Content, nodeId, depth, template: true);
        }

        var frames = document.QuerySelectorAll("iframe")
            .OfType<IHtmlInlineFrameElement>()
            .Where(frameNodeIds.ContainsKey)
            .Select(frame => (Element: frame, NodeId: frameNodeIds[frame]))
            .ToArray();
        if (frameDepth >= MaxFrameDepth) {
            foreach ((IHtmlInlineFrameElement frame, _) in frames) {
                IDocument? child = frame.ContentDocument;
                if (child != null && CanCaptureFrame(documentUrl, child, out _)) {
                    throw new HtmlScriptRuntimeException("Captured frame depth budget exceeded.");
                }
            }
            return result;
        }
        foreach ((IHtmlInlineFrameElement frame, int nodeId) in frames) {
            token.ThrowIfCancellationRequested();
            IDocument? child = frame.ContentDocument;
            if (child == null || !CanCaptureFrame(documentUrl, child, out Uri childUrl)) continue;
            HtmlRuntimeWireDocument captured = CaptureDocument(child, request, token, budget, frameDepth + 1, childUrl);
            captured.DocumentUrl = childUrl;
            captured.BaseUri = ResolveBaseUri(child, childUrl);
            result.Frames.Add(new HtmlRuntimeWireFrame { FrameElementNodeId = nodeId, Document = captured });
        }
        return result;
    }

    private static bool CanCaptureFrame(
        Uri containingUrl,
        IDocument child,
        out Uri childUrl) {
        childUrl = ResolveDocumentUrl(child, containingUrl);
        if ((child.Context.Security & Sandboxes.Origin) == Sandboxes.Origin) return false;
        return string.Equals(
            containingUrl.GetLeftPart(UriPartial.Authority),
            childUrl.GetLeftPart(UriPartial.Authority),
            StringComparison.OrdinalIgnoreCase);
    }

    private static Uri ResolveDocumentUrl(IDocument document, Uri? inherited) {
        if (Uri.TryCreate(document.Url, UriKind.Absolute, out Uri? parsed)
            && (parsed.Scheme == Uri.UriSchemeHttp || parsed.Scheme == Uri.UriSchemeHttps)) {
            return HtmlRuntimeResourcePolicy.ValidateUrl(parsed);
        }
        if (inherited != null) return inherited;
        throw new HtmlScriptRuntimeException("Captured documents require an HTTP(S) identity.");
    }

    private static Uri ResolveBaseUri(IDocument document, Uri documentUrl) {
        string value = RuntimeDocumentUrls.Base(document);
        return Uri.TryCreate(value, UriKind.Absolute, out Uri? parsed)
            && parsed.AbsoluteUri.Length <= 8192
            && parsed.Scheme is not ("data" or "javascript")
                ? parsed
                : documentUrl;
    }

    private sealed class CaptureBudget {
        private readonly HtmlScriptRequest _request;
        private int _nodes;
        private long _characters;

        internal CaptureBudget(HtmlScriptRequest request) => _request = request;

        internal void ReserveNode() {
            if (++_nodes > _request.MaxNodes) throw new HtmlScriptRuntimeException("Captured node budget exceeded.");
        }

        internal void Count(string? value) {
            _characters += value?.Length ?? 0;
            if (_characters > _request.MaxOutputCharacters)
                throw new HtmlScriptRuntimeException("Captured data budget exceeded.");
        }
    }
}
