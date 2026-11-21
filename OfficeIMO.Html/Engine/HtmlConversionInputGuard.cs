using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Applies shared source and DOM limits before expensive conversion analysis begins.</summary>
internal static partial class HtmlConversionInputGuard {
    internal const int MaxSrcDocDepth = 8;

    internal static void ValidateSource(string html, HtmlConversionLimits limits) {
        if (!limits.MaxInputCharacters.HasValue || html.Length <= limits.MaxInputCharacters.Value) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded,
            "HTML source length exceeded the configured conversion limit.",
            nameof(HtmlConversionLimits.MaxInputCharacters),
            html.Length,
            limits.MaxInputCharacters.Value);
    }

    internal static int ValidateDocument(IDocument document, HtmlConversionLimits limits, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlDomLimitTracker? tracker = HtmlDomLimitTracker.Create(limits.MaxHtmlNodes, limits.MaxHtmlDepth);
        var cssBudget = new HtmlCssByteBudget(limits);
        int nodes = 0;

        // Keep one position per ancestor rather than buffering every sibling before checking its
        // budget. Child lists are read by index, and a node without children adds no level.
        var pending = new PendingNodes();
        pending.Push(document.ChildNodes, 0, 0);
        try {
            while (pending.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!pending.TryTakeNext(out INode node, out int parentDepth, out int srcDocDepth)) {
                    pending.Pop();
                    continue;
                }

                nodes = checked(nodes + 1);
                int depth = parentDepth + (node is IElement ? 1 : 0);
                if (node is IElement element) {
                    tracker?.RecordElementStart(depth);
                    ValidateSemanticAttributes(element, limits.MaxSemanticMetadataCharacters, cancellationToken);
                    if (string.Equals(element.LocalName, "style", StringComparison.OrdinalIgnoreCase)) {
                        cssBudget.ReserveOrThrow(HtmlStyleElementText.Read(element) ?? string.Empty);
                    }
                    string? inlineStyle = element.GetAttribute("style");
                    if (!string.IsNullOrEmpty(inlineStyle)) cssBudget.ReserveOrThrow(inlineStyle!);
                    if (srcDocDepth < MaxSrcDocDepth) {
                        string? source = element.GetAttribute("srcdoc");
                        if (!string.IsNullOrWhiteSpace(source)) {
                            ValidateSource(source!, limits);
                            IHtmlDocument nested = HtmlDocumentParser.ParseDocument(source!, cancellationToken);
                            pending.Push(nested.ChildNodes, depth, srcDocDepth + 1);
                        }
                    }
                    if (element is IHtmlTemplateElement template) {
                        pending.PushSingle(template.Content, depth, srcDocDepth);
                    }
                } else {
                    tracker?.RecordNode();
                }
                pending.Push(node.ChildNodes, depth, srcDocDepth);
            }
        } finally {
            pending.Clear();
        }
        cancellationToken.ThrowIfCancellationRequested();
        return nodes;
    }

    // The unvisited rest of each ancestor's child list, innermost last.
    private sealed class PendingNodes {
        private Level[] _levels = new Level[16];

        internal int Count { get; private set; }

        internal void Push(INodeList children, int parentDepth, int srcDocDepth) {
            if (children.Length == 0) {
                return;
            }

            Add(new Level(children, null, parentDepth, srcDocDepth));
        }

        internal void PushSingle(INode node, int parentDepth, int srcDocDepth) =>
            Add(new Level(null, node, parentDepth, srcDocDepth));

        private void Add(Level level) {
            if (Count == _levels.Length) {
                Array.Resize(ref _levels, _levels.Length * 2);
            }

            _levels[Count++] = level;
        }

        internal bool TryTakeNext(out INode node, out int parentDepth, out int srcDocDepth) {
            ref Level level = ref _levels[Count - 1];
            parentDepth = level.ParentDepth;
            srcDocDepth = level.SrcDocDepth;
            if (level.Single != null) {
                node = level.Single;
                level.Single = null;
                level.Index = 1;
                return true;
            }

            if (level.Children != null && level.Index < level.Children.Length) {
                node = level.Children[level.Index++];
                return true;
            }

            node = null!;
            return false;
        }

        internal void Pop() => _levels[--Count] = default;

        internal void Clear() {
            Array.Clear(_levels, 0, Count);
            Count = 0;
        }

        private struct Level {
            internal Level(INodeList? children, INode? single, int parentDepth, int srcDocDepth) {
                Children = children;
                Single = single;
                Index = 0;
                ParentDepth = parentDepth;
                SrcDocDepth = srcDocDepth;
            }

            internal INodeList? Children;
            internal INode? Single;
            internal int Index;
            internal int ParentDepth;
            internal int SrcDocDepth;
        }
    }
    private static void ValidateSemanticAttributes(IElement element, int? maximumCharacters, CancellationToken cancellationToken) {
        if (!maximumCharacters.HasValue) return;
        var attributes = element.Attributes;
        for (int index = 0; index < attributes.Length; index++) {
            IAttr attribute = attributes[index]!;
            cancellationToken.ThrowIfCancellationRequested();
            if (!attribute.Name.StartsWith("data-officeimo-", StringComparison.OrdinalIgnoreCase)
                || attribute.Value.Length <= maximumCharacters.Value) {
                continue;
            }

            throw new HtmlDomLimitException(
                HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "OfficeIMO semantic metadata exceeded the configured conversion limit.",
                nameof(HtmlConversionLimits.MaxSemanticMetadataCharacters),
                attribute.Value.Length,
                maximumCharacters.Value);
        }
    }

}
