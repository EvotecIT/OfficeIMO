using System.Runtime.CompilerServices;
using System.Threading;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using HtmlElement = OfficeIMO.Html.Dom.HtmlElement;

namespace OfficeIMO.Html;

/// <summary>Explicit structural bridge for the existing native-DOM CSS/layout implementation. Never reparses serialized source.</summary>
internal static class NativeDomBridge {
    private static readonly ConditionalWeakTable<HtmlDocument, NativeState> OwnedStates = new ConditionalWeakTable<HtmlDocument, NativeState>();
    private static readonly ConditionalWeakTable<IHtmlDocument, CallbackState> CallbackSnapshots = new ConditionalWeakTable<IHtmlDocument, CallbackState>();
    private static readonly ConditionalWeakTable<HtmlDocument, CallbackState> CallbackSources = new ConditionalWeakTable<HtmlDocument, CallbackState>();
    private static readonly object CacheSync = new object();

    internal sealed class NativeState {
        internal NativeState(IHtmlDocument native, HtmlDocument owned) { Native = native; Owned = owned; Revision = owned.Revision; }
        internal IHtmlDocument Native { get; }
        internal HtmlDocument Owned { get; }
        internal long Revision { get; set; }
        internal Dictionary<int, INode> ToNative { get; } = new Dictionary<int, INode>();
        internal Dictionary<INode, HtmlNode> ToOwned { get; } = new Dictionary<INode, HtmlNode>();
        internal void Add(INode native, HtmlNode owned) { ToNative.Add(owned.NodeId, native); ToOwned.Add(native, owned); }
    }

    internal static HtmlDocument Import(IHtmlDocument native, HtmlParseOptions? options = null, CancellationToken cancellationToken = default) {
        if (native == null) throw new ArgumentNullException(nameof(native));
        var mode = ((AngleSharp.Html.Construction.IConstructableDocument)native).QuirksMode;
        var owned = new HtmlDocument(AngleSharpDomServices.Instance, AngleSharpHtmlParser.Instance.Id,
            mode == QuirksMode.On ? HtmlDocumentMode.Quirks : mode == QuirksMode.Limited ? HtmlDocumentMode.LimitedQuirks : HtmlDocumentMode.Standards);
        var state = new NativeState(native, owned);
        state.Add(native, owned);
        var pending = new Stack<(INode Native, HtmlNode Owned, int Depth)>();
        pending.Push((native, owned, 0));
        int nodes = 0;
        while (pending.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var current = pending.Pop();
            foreach (INode child in current.Native.ChildNodes) {
                int depth = current.Depth + (child is IElement ? 1 : 0);
                if (options?.MaxNodes is int maxNodes && ++nodes > maxNodes) throw new HtmlParseLimitException(nameof(options.MaxNodes), nodes, maxNodes);
                if (options?.MaxDepth is int maxDepth && child is IElement && depth > maxDepth) throw new HtmlParseLimitException(nameof(options.MaxDepth), depth, maxDepth);
                HtmlNode converted = ImportNode(child, owned);
                current.Owned.AppendChild(converted);
                state.Add(child, converted);
                pending.Push((child, converted, depth));
            }
            if (current.Native is IHtmlTemplateElement template && current.Owned is HtmlElement element) {
                if (options?.MaxNodes is int maxNodes && ++nodes > maxNodes) throw new HtmlParseLimitException(nameof(options.MaxNodes), nodes, maxNodes);
                HtmlNode content = owned.GetOrCreateTemplateContent(element);
                state.Add(template.Content, content);
                pending.Push((template.Content, content, current.Depth));
            }
        }
        state.Revision = owned.Revision;
        lock (CacheSync) {
            OwnedStates.Add(owned, state);
        }
        return owned;
    }

    private sealed class CallbackState {
        internal Dictionary<INode, HtmlNode> ToOwned { get; } = new Dictionary<INode, HtmlNode>();
        internal Dictionary<int, INode> ToOriginal { get; } = new Dictionary<int, INode>();
    }

    internal static HtmlElement Wrap(IElement native) {
        IHtmlDocument document = native.Owner as IHtmlDocument ?? throw new ArgumentException("An HTML document is required.", nameof(native));
        lock (CacheSync) {
            if (!CallbackSnapshots.TryGetValue(document, out CallbackState? state)) {
                IHtmlDocument clone = HtmlDocumentParser.CloneDocument(document);
                HtmlDocument owned = Import(clone).Freeze();
                NativeState clonedState = GetState(owned);
                state = new CallbackState();
                var pending = new Stack<(INode Original, INode Clone)>();
                pending.Push((document, clone));
                while (pending.Count != 0) {
                    var pair = pending.Pop();
                    HtmlNode ownedNode = clonedState.ToOwned[pair.Clone];
                    state.ToOwned.Add(pair.Original, ownedNode);
                    state.ToOriginal.Add(ownedNode.NodeId, pair.Original);
                    for (int index = 0; index < pair.Original.ChildNodes.Length; index++) pending.Push((pair.Original.ChildNodes[index], pair.Clone.ChildNodes[index]));
                    if (pair.Original is IHtmlTemplateElement sourceTemplate && pair.Clone is IHtmlTemplateElement cloneTemplate) pending.Push((sourceTemplate.Content, cloneTemplate.Content));
                }
                CallbackSnapshots.Add(document, state);
                CallbackSources.Add(owned, state);
            }
            if (!state.ToOwned.TryGetValue(native, out HtmlNode? node)) throw new InvalidOperationException("The element is outside the retained callback snapshot.");
            return (HtmlElement)node;
        }
    }

    internal static INode GetCallbackNative(HtmlNode node, HtmlDocument scope) {
        if (node == null || !ReferenceEquals(node.Document, scope)) throw new ArgumentException("Callback nodes must belong to the current element's snapshot.", nameof(node));
        lock (CacheSync) {
            if (!CallbackSources.TryGetValue(scope, out CallbackState? state) || !state.ToOriginal.TryGetValue(node.NodeId, out INode? native)) throw new ArgumentException("The node is not part of this conversion callback.", nameof(node));
            return native;
        }
    }

    internal static void ReleaseCallbackSnapshot(IHtmlDocument document) {
        lock (CacheSync) CallbackSnapshots.Remove(document);
    }

    internal static IHtmlDocument GetNativeDocument(HtmlDocument document) => GetState(document).Native;
    internal static INode GetNative(HtmlNode node) {
        if (node == null) throw new ArgumentNullException(nameof(node));
        return GetState(node.Document).ToNative[node.NodeId];
    }

    internal static NativeState GetState(HtmlDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        lock (CacheSync) {
            if (OwnedStates.TryGetValue(document, out NativeState? state) && state.Revision == document.Revision) return state;
            state = Export(document);
            OwnedStates.Remove(document);
            OwnedStates.Add(document, state);
            return state;
        }
    }

    private static HtmlNode ImportNode(INode source, HtmlDocument document) {
        if (source is IElement element) {
            HtmlElement result = document.CreateElement(element.LocalName, element.NamespaceUri ?? string.Empty, element.Prefix);
            foreach (IAttr attribute in element.Attributes) result.SetAttribute(attribute.Name, attribute.Value, attribute.NamespaceUri);
            if (element.SourceReference?.Position.Index is int index && index >= 0) document.SetSourceIndex(result, index);
            return result;
        }
        if (source is IDocumentType type) return document.CreateDocumentType(type.Name, type.PublicIdentifier, type.SystemIdentifier);
        if (source.NodeType == NodeType.Comment) return document.CreateComment(source.TextContent ?? string.Empty);
        if (source.NodeType == NodeType.DocumentFragment) return document.CreateFragment();
        if (source.NodeType == NodeType.Text) return document.CreateTextNode(source.TextContent ?? string.Empty);
        throw new NotSupportedException("The HTML provider cannot import node kind " + source.NodeType + ".");
    }

    private static NativeState Export(HtmlDocument document) {
        // Parse only a constant empty envelope to choose the existing DOM implementation and mode.
        // The caller's content is materialized structurally, never serialized and reparsed.
        IHtmlDocument native = HtmlDocumentParser.ParseDocument("<!doctype html>");
        ((AngleSharp.Html.Construction.IConstructableDocument)native).QuirksMode = document.Mode == HtmlDocumentMode.Quirks ? QuirksMode.On : document.Mode == HtmlDocumentMode.LimitedQuirks ? QuirksMode.Limited : QuirksMode.Off;
        foreach (INode child in native.ChildNodes.ToArray()) native.RemoveChild(child);
        var state = new NativeState(native, document);
        state.Add(native, document);
        var pending = new Stack<(HtmlNode Owned, INode Native)>();
        pending.Push((document, native));
        foreach (HtmlNode detached in document.Nodes.Where(node => node.Parent == null && node.TemplateHost == null && node.Kind != HtmlNodeKind.Document)) {
            INode converted = ExportNode(detached, native);
            state.Add(converted, detached);
            pending.Push((detached, converted));
        }
        while (pending.Count != 0) {
            var current = pending.Pop();
            foreach (HtmlNode child in current.Owned.ChildNodes) {
                INode converted = ExportNode(child, native);
                current.Native.AppendChild(converted);
                state.Add(converted, child);
                pending.Push((child, converted));
            }
            if (current.Owned is HtmlElement element && element.TemplateContent != null && current.Native is IHtmlTemplateElement template) {
                state.Add(template.Content, element.TemplateContent);
                pending.Push((element.TemplateContent, template.Content));
            }
        }
        return state;
    }

    private static INode ExportNode(HtmlNode source, IHtmlDocument document) {
        if (source is HtmlElement element) {
            // Parser construction accepts recovered HTML names that the XML-name authoring API rejects.
            // Use the same factories structurally so editing another node cannot invalidate that recovery.
            var owner = (AngleSharp.Dom.Document)document;
            string? prefix = element.Prefix.Length == 0 ? null : element.Prefix;
            IElement result = element.NamespaceUri == HtmlElement.HtmlNamespace
                ? owner.CreateElementFrom(element.LocalName, prefix!, NodeFlags.None)
                : element.NamespaceUri == "http://www.w3.org/2000/svg"
                    ? owner.Context.GetService<IElementFactory<AngleSharp.Dom.Document, AngleSharp.Svg.Dom.SvgElement>>()!.Create(owner, element.LocalName, prefix, NodeFlags.None)
                    : element.NamespaceUri == "http://www.w3.org/1998/Math/MathML"
                        ? owner.Context.GetService<IElementFactory<AngleSharp.Dom.Document, AngleSharp.Mathml.Dom.MathElement>>()!.Create(owner, element.LocalName, prefix, NodeFlags.None)
                        : document.CreateElement(element.NamespaceUri, element.NodeName);
            var constructable = (AngleSharp.Html.Construction.IConstructableElement)result;
            foreach (HtmlAttribute attribute in element.Attributes) {
                if (attribute.NamespaceUri.Length == 0) constructable.SetOwnAttribute(attribute.Name, attribute.Value);
                else constructable.SetAttribute(attribute.NamespaceUri, attribute.Name, attribute.Value);
            }
            return result;
        }
        if (source is HtmlDocumentType type) {
            var owner = (AngleSharp.Dom.Document)document;
            var factory = owner.Context.GetService<AngleSharp.Html.Construction.IHtmlElementConstructionFactory>()!;
            return (INode)factory.CreateDocumentType(owner, type.Name, type.PublicIdentifier, type.SystemIdentifier);
        }
        if (source.Kind == HtmlNodeKind.Text) return document.CreateTextNode(source.TextContent);
        if (source.Kind == HtmlNodeKind.Comment) return document.CreateComment(source.TextContent);
        if (source.Kind == HtmlNodeKind.DocumentFragment) return document.CreateDocumentFragment();
        throw new NotSupportedException("The HTML provider cannot export node kind " + source.Kind + ".");
    }
}
