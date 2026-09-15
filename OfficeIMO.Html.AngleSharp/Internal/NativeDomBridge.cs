using System.Collections.Concurrent;
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
    // The owned document is the public source of truth. Keep its replaceable native projection
    // reclaimable so a long-lived owned snapshot does not permanently retain a second tree and
    // two node maps. A caller actively holding the native document can keep that exact projection;
    // the bridge can rebuild its maps without keeping either graph alive by itself.
    private static readonly ConditionalWeakTable<HtmlDocument, NativeStateReference> OwnedStates = new ConditionalWeakTable<HtmlDocument, NativeStateReference>();
    private static readonly ConditionalWeakTable<IHtmlDocument, Lazy<CallbackState>> CallbackSnapshots = new ConditionalWeakTable<IHtmlDocument, Lazy<CallbackState>>();
    private static readonly ConditionalWeakTable<HtmlDocument, CallbackState> CallbackSources = new ConditionalWeakTable<HtmlDocument, CallbackState>();
    private static readonly object CacheSync = new object();
    private sealed class NativeStateReference {
        private readonly WeakReference<NativeState> _state;
        private readonly WeakReference<IHtmlDocument> _native;
        private readonly long _revision;
        internal NativeStateReference(NativeState state) {
            _state = new WeakReference<NativeState>(state);
            _native = new WeakReference<IHtmlDocument>(state.Native);
            _revision = state.Revision;
        }
        internal bool TryGet(out NativeState? state) => _state.TryGetTarget(out state);
        internal bool TryGetNative(long revision, out IHtmlDocument? native) {
            native = null;
            return revision == _revision && _native.TryGetTarget(out native);
        }
    }

    internal sealed class NativeState {
        internal NativeState(IHtmlDocument native, HtmlDocument owned) { Native = native; Owned = owned; Revision = owned.Revision; }
        internal IHtmlDocument Native { get; }
        internal HtmlDocument Owned { get; }
        internal long Revision { get; set; }
        internal object DetachedSync { get; } = new object();
        // Frozen snapshots can materialize separate detached roots concurrently with attached-tree reads.
        internal ConcurrentDictionary<int, INode> ToNative { get; } = new ConcurrentDictionary<int, INode>();
        internal ConcurrentDictionary<INode, HtmlNode> ToOwned { get; } = new ConcurrentDictionary<INode, HtmlNode>();
        internal void Add(INode native, HtmlNode owned) {
            if (!ToNative.TryAdd(owned.NodeId, native) || !ToOwned.TryAdd(native, owned)) throw new InvalidOperationException("A node has already been mapped.");
        }
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
                cancellationToken.ThrowIfCancellationRequested();
                int depth = current.Depth + (child is IElement ? 1 : 0);
                if (options?.MaxNodes is int maxNodes && ++nodes > maxNodes) throw new HtmlParseLimitException(nameof(options.MaxNodes), nodes, maxNodes);
                if (options?.MaxDepth is int maxDepth && child is IElement && depth > maxDepth) throw new HtmlParseLimitException(nameof(options.MaxDepth), depth, maxDepth);
                HtmlNode converted = ImportNode(child, owned, cancellationToken);
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
        cancellationToken.ThrowIfCancellationRequested();
        lock (CacheSync) {
            cancellationToken.ThrowIfCancellationRequested();
            OwnedStates.Add(owned, new NativeStateReference(state));
        }
        return owned;
    }

    internal static HtmlDocumentFragment ImportFragment(
        IEnumerable<INode> nativeNodes,
        HtmlDocumentMode mode,
        HtmlParseOptions? options = null,
        CancellationToken cancellationToken = default,
        int sourceIndexOffset = 0) {
        if (nativeNodes == null) throw new ArgumentNullException(nameof(nativeNodes));
        var owned = new HtmlDocument(AngleSharpDomServices.Instance, AngleSharpHtmlParser.Instance.Id, mode);
        HtmlDocumentFragment fragment = owned.CreateFragment();
        var pending = new Stack<(INode Native, HtmlNode Owned, int Depth)>();
        int nodes = 0;
        foreach (INode native in nativeNodes) {
            cancellationToken.ThrowIfCancellationRequested();
            int depth = native is IElement ? 1 : 0;
            RecordImportedNode(options, ref nodes, depth, native is IElement);
            HtmlNode converted = ImportNode(native, owned, cancellationToken, sourceIndexOffset);
            fragment.AppendChild(converted);
            pending.Push((native, converted, depth));
        }
        while (pending.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var current = pending.Pop();
            foreach (INode child in current.Native.ChildNodes) {
                cancellationToken.ThrowIfCancellationRequested();
                int depth = current.Depth + (child is IElement ? 1 : 0);
                RecordImportedNode(options, ref nodes, depth, child is IElement);
                HtmlNode converted = ImportNode(child, owned, cancellationToken, sourceIndexOffset);
                current.Owned.AppendChild(converted);
                pending.Push((child, converted, depth));
            }
            if (current.Native is IHtmlTemplateElement template && current.Owned is HtmlElement element) {
                RecordImportedNode(options, ref nodes, current.Depth, isElement: false);
                HtmlDocumentFragment content = owned.GetOrCreateTemplateContent(element);
                pending.Push((template.Content, content, current.Depth));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        owned.Freeze();
        return fragment;
    }

    private static void RecordImportedNode(HtmlParseOptions? options, ref int nodes, int depth, bool isElement) {
        nodes++;
        if (options?.MaxNodes is int maxNodes && nodes > maxNodes)
            throw new HtmlParseLimitException(nameof(options.MaxNodes), nodes, maxNodes);
        if (isElement && options?.MaxDepth is int maxDepth && depth > maxDepth)
            throw new HtmlParseLimitException(nameof(options.MaxDepth), depth, maxDepth);
    }

    private sealed class CallbackState {
        internal Dictionary<INode, HtmlNode> ToOwned { get; } = new Dictionary<INode, HtmlNode>();
        internal Dictionary<int, INode> ToOriginal { get; } = new Dictionary<int, INode>();
    }

    internal static HtmlElement Wrap(IElement native) {
        IHtmlDocument document = native.Owner as IHtmlDocument ?? throw new ArgumentException("An HTML document is required.", nameof(native));
        // Only publication is global. Snapshot construction is serialized for this native document.
        Lazy<CallbackState> snapshot;
        lock (CacheSync) {
            if (!CallbackSnapshots.TryGetValue(document, out snapshot!)) {
                snapshot = new Lazy<CallbackState>(() => CreateCallbackSnapshot(document), LazyThreadSafetyMode.ExecutionAndPublication);
                CallbackSnapshots.Add(document, snapshot);
            }
        }
        CallbackState state;
        try { state = snapshot.Value; }
        catch {
            lock (CacheSync) {
                if (CallbackSnapshots.TryGetValue(document, out Lazy<CallbackState>? current) && ReferenceEquals(current, snapshot)) CallbackSnapshots.Remove(document);
            }
            throw;
        }
        if (!state.ToOwned.TryGetValue(native, out HtmlNode? node)) throw new InvalidOperationException("The element is outside the retained callback snapshot.");
        return (HtmlElement)node;
    }

    private static CallbackState CreateCallbackSnapshot(IHtmlDocument document) {
        IHtmlDocument clone = HtmlDocumentParser.CloneDocument(document);
        HtmlDocument owned = Import(clone).Freeze();
        NativeState clonedState = GetState(owned);
        var state = new CallbackState();
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
        lock (CacheSync) CallbackSources.Add(owned, state);
        return state;
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

    internal static IHtmlDocument GetNativeDocument(HtmlDocument document, CancellationToken cancellationToken = default) => GetState(document, cancellationToken).Native;
    /// <summary>Preserves optional-node inputs for helpers whose established contract handles missing elements.</summary>
    internal static INode? GetNativeOrNull(HtmlNode? node) => node == null ? null : GetNative(node);
    internal static INode GetNative(HtmlNode node) {
        if (node == null) throw new ArgumentNullException(nameof(node));
        return GetState(node).ToNative[node.NodeId];
    }

    internal static NativeState GetState(HtmlNode node) {
        if (node == null) throw new ArgumentNullException(nameof(node));
        NativeState state = GetState(node.Document);
        lock (state.DetachedSync) {
            if (!state.ToNative.ContainsKey(node.NodeId)) {
                HtmlNode root = node;
                while (root.Parent != null || root.TemplateHost != null) root = root.Parent ?? root.TemplateHost!;
                // Detached trees are exported only when directly queried or serialized. Build mappings
                // separately so an invalid detached tree cannot poison the attached conversion cache.
                var detached = new NativeState(state.Native, node.Document);
                INode native = ExportNode(root, state.Native);
                detached.Add(native, root);
                ExportChildren(detached, root, native);
                foreach (var pair in detached.ToOwned) state.Add(pair.Key, pair.Value);
            }
            return state;
        }
    }

    internal static NativeState GetState(HtmlDocument document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        IHtmlDocument? retainedNative = null;
        lock (CacheSync) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryGetCurrentState(document, out NativeState? state)) return state!;
            if (OwnedStates.TryGetValue(document, out NativeStateReference? reference))
                reference.TryGetNative(document.Revision, out retainedNative);
        }
        // Concurrent first readers may build candidates, but only one complete state is installed.
        // No tree traversal or provider materialization may hold the process-wide cache lock.
        NativeState candidate = retainedNative == null
            ? Export(document, cancellationToken)
            : RebuildState(retainedNative, document, cancellationToken) ?? Export(document, cancellationToken);
        lock (CacheSync) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryGetCurrentState(document, out NativeState? state)) return state!;
            OwnedStates.Remove(document);
            OwnedStates.Add(document, new NativeStateReference(candidate));
            return candidate;
        }
    }

    private static bool TryGetCurrentState(HtmlDocument document, out NativeState? state) {
        state = null;
        return OwnedStates.TryGetValue(document, out NativeStateReference? reference)
            && reference.TryGet(out state)
            && state != null
            && state.Revision == document.Revision;
    }

    private static NativeState? RebuildState(IHtmlDocument native, HtmlDocument owned, CancellationToken cancellationToken) {
        var state = new NativeState(native, owned);
        var pending = new Stack<(INode Native, HtmlNode Owned)>();
        pending.Push((native, owned));
        while (pending.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var current = pending.Pop();
            if (!NodesMatch(current.Native, current.Owned)
                || current.Native.ChildNodes.Length != current.Owned.ChildNodes.Count) return null;
            state.Add(current.Native, current.Owned);
            for (int index = 0; index < current.Native.ChildNodes.Length; index++)
                pending.Push((current.Native.ChildNodes[index], current.Owned.ChildNodes[index]));
            if (current.Native is IHtmlTemplateElement nativeTemplate && current.Owned is HtmlElement ownedTemplate) {
                if (ownedTemplate.TemplateContent == null) return null;
                pending.Push((nativeTemplate.Content, ownedTemplate.TemplateContent));
            } else if (current.Owned is HtmlElement { TemplateContent: not null }) return null;
        }
        return state;
    }

    private static bool NodesMatch(INode native, HtmlNode owned) {
        if (native is IElement nativeElement && owned is HtmlElement ownedElement)
            return string.Equals(nativeElement.LocalName, ownedElement.LocalName, StringComparison.Ordinal)
                && string.Equals(nativeElement.NamespaceUri ?? string.Empty, ownedElement.NamespaceUri, StringComparison.Ordinal);
        return native.NodeType switch {
            NodeType.Document => owned.Kind == HtmlNodeKind.Document,
            NodeType.DocumentFragment => owned.Kind == HtmlNodeKind.DocumentFragment,
            NodeType.DocumentType => owned.Kind == HtmlNodeKind.DocumentType,
            NodeType.Text => owned.Kind == HtmlNodeKind.Text,
            NodeType.Comment => owned.Kind == HtmlNodeKind.Comment,
            _ => false
        };
    }

    private static HtmlNode ImportNode(INode source, HtmlDocument document, CancellationToken cancellationToken, int sourceIndexOffset = 0) {
        if (source is IElement element) {
            HtmlElement result = document.CreateElement(element.LocalName, element.NamespaceUri ?? string.Empty, element.Prefix);
            foreach (IAttr attribute in element.Attributes) {
                cancellationToken.ThrowIfCancellationRequested();
                result.SetAttribute(new HtmlAttribute(attribute.Name, attribute.Value, attribute.NamespaceUri));
            }
            if (element.SourceReference?.Position.Index is int index && index >= sourceIndexOffset)
                document.SetSourceIndex(result, index - sourceIndexOffset);
            result.FormState = NativeFormState.Get(element);
            return result;
        }
        if (source is IDocumentType type) return document.CreateDocumentType(type.Name, type.PublicIdentifier, type.SystemIdentifier);
        if (source.NodeType == NodeType.Comment) return document.CreateComment(source.TextContent ?? string.Empty);
        if (source.NodeType == NodeType.DocumentFragment) return document.CreateFragment();
        if (source.NodeType == NodeType.Text) return document.CreateTextNode(source.TextContent ?? string.Empty);
        throw new NotSupportedException("The HTML provider cannot import node kind " + source.NodeType + ".");
    }

    private static NativeState Export(HtmlDocument document, CancellationToken cancellationToken) {
        // Parse only a constant empty envelope to choose the existing DOM implementation and mode.
        // The caller's content is materialized structurally, never serialized and reparsed.
        IHtmlDocument native = HtmlDocumentParser.ParseDocument("<!doctype html>", cancellationToken);
        ((AngleSharp.Html.Construction.IConstructableDocument)native).QuirksMode = document.Mode == HtmlDocumentMode.Quirks ? QuirksMode.On : document.Mode == HtmlDocumentMode.LimitedQuirks ? QuirksMode.Limited : QuirksMode.Off;
        foreach (INode child in native.ChildNodes.ToArray()) native.RemoveChild(child);
        var state = new NativeState(native, document);
        state.Add(native, document);
        ExportChildren(state, document, native, cancellationToken);
        return state;
    }

    private static void ExportChildren(NativeState state, HtmlNode root, INode nativeRoot, CancellationToken cancellationToken = default) {
        var pending = new Stack<(HtmlNode Owned, INode Native)>();
        pending.Push((root, nativeRoot));
        while (pending.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var current = pending.Pop();
            foreach (HtmlNode child in current.Owned.ChildNodes) {
                cancellationToken.ThrowIfCancellationRequested();
                INode converted = ExportNode(child, state.Native, cancellationToken);
                current.Native.AppendChild(converted);
                state.Add(converted, child);
                pending.Push((child, converted));
            }
            if (current.Owned is HtmlElement element && element.TemplateContent != null && current.Native is IHtmlTemplateElement template) {
                state.Add(template.Content, element.TemplateContent);
                pending.Push((element.TemplateContent, template.Content));
            }
        }
        NativeFormState.ApplyTree(nativeRoot, cancellationToken);
    }

    private static INode ExportNode(HtmlNode source, IHtmlDocument document, CancellationToken cancellationToken = default) {
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
                cancellationToken.ThrowIfCancellationRequested();
                if (attribute.NamespaceUri.Length == 0) constructable.SetOwnAttribute(attribute.Name, attribute.Value);
                else constructable.SetAttribute(attribute.NamespaceUri, attribute.Name, attribute.Value);
            }
            // Parser factories defer initialization until attributes have been supplied.
            // Complete that lifecycle before exposing the node to clone/layout consumers.
            constructable.SetupElement();
            NativeFormState.Attach(result, element.FormState);
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
