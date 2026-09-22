using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

// Used only under the session's realm lock. Nested synchronous calls retain the
// outer entry document, including calls through another realm's DOM wrappers.
internal sealed class RuntimeScriptEntry {
    internal IDocument? Document { get; private set; }

    internal IDisposable Enter(IDocument? document) {
        var prior = Document;
        Document ??= document;
        return new Scope(this, prior);
    }

    private sealed class Scope(RuntimeScriptEntry owner, IDocument? prior) : IDisposable {
        public void Dispose() => owner.Document = prior;
    }
}
