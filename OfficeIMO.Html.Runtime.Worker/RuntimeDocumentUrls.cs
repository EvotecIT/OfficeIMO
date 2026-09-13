using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeDocumentUrls {
    private sealed class State {
        internal IHtmlBaseElement? Element;
        internal string? Href;
        internal Url? Url;
        internal bool Invalidated;
        internal void Invalidate(IMutationRecord record) {
            if(record.Type=="attributes" && record.AttributeName=="href" && string.IsNullOrEmpty(record.AttributeNamespace) && record.Target is IHtmlBaseElement changed && Element!=null) {
                var order=changed.CompareDocumentPosition(Element);
                if(ReferenceEquals(changed,Element) || (order & DocumentPositions.Disconnected)==0 && (order & DocumentPositions.Following)!=0)
                    Invalidated=true;
            }
            if(record.Type!="childList" || Element==null)return;
            foreach(var nodes in new[]{record.Added,record.Removed}) {
                if(nodes==null)continue;
                foreach(var node in nodes) {
                    for(INode? ancestor=Element;ancestor!=null;ancestor=ancestor.Parent)
                        if(ReferenceEquals(node,ancestor))Invalidated=true;
                    // A temporary earlier base can activate and disappear before
                    // the next URL read. Its recorded insertion/removal position
                    // distinguishes it from an unrelated later base.
                    bool containsBase=node is IHtmlBaseElement candidate && candidate.HasAttribute("href")
                        || node is IParentNode parent && parent.QuerySelectorAll("base[href]").Any(child=>child is IHtmlBaseElement);
                    var previous=record.PreviousSibling ?? record.Target;
                    var position=previous.CompareDocumentPosition(Element);
                    if(containsBase && (position & DocumentPositions.Disconnected)==0 && (position & DocumentPositions.Following)!=0)Invalidated=true;
                }
                }
        }
    }
    private static readonly ConditionalWeakTable<IDocument, State> States = new();

    internal sealed class MutationListener : IDomMutationListener {
        public void OnMutation(IDocument document,IMutationRecord record) {
            if(States.TryGetValue(document,out var state))state.Invalidate(record);
        }
    }

    internal static string Base(IDocument document) {
        var state = States.GetOrCreateValue(document);
        var element = document.QuerySelectorAll("base[href]").OfType<IHtmlBaseElement>().FirstOrDefault();
        string? href = element?.GetAttribute("href");
        if (state.Invalidated || !ReferenceEquals(element,state.Element) || href != state.Href) {
            state.Invalidated=false;
            state.Element=element;state.Href=href;
            var candidate=href == null ? null : new Url(new Url(document.Url),href);
            state.Url=candidate is {IsInvalid:false} && candidate.Scheme is not ("data" or "javascript") ? candidate : null;
        }
        // Freeze a base element's resolved URL until its href or active identity
        // changes. History rewrites preserve that base, while no-base pages follow URL.
        if (document is Document native) native.BaseUrl=state.Url;
        return state.Url?.Href ?? document.Url;
    }

    internal static void Rewrite(IDocument document, string url) {
        Base(document);
        var native=(Document)document;
        var parsed=new Url(url);
        native.DocumentUrl.Href=parsed.Href;
        // The retained URL setter reparses relative to its old record and leaves
        // its fragment behind when the new serialized URL has none.
        native.DocumentUrl.Fragment=parsed.Fragment;
        native.DocumentUrl.Query=parsed.Query;
    }
}
