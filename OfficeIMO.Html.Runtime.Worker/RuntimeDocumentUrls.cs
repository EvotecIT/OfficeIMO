using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeDocumentUrls {
    internal static string Base(IDocument document) => document.BaseUri;

    internal static string Origin(IDocument document) => document.Origin ?? "null";

    internal static void Rewrite(IDocument document, string url) {
        var native=(Document)document;
        var parsed=new Url(url);
        native.DocumentUrl.Href=parsed.Href;
        // The retained URL setter reparses relative to its old record and leaves
        // its fragment behind when the new serialized URL has none.
        native.DocumentUrl.Fragment=parsed.Fragment;
        native.DocumentUrl.Query=parsed.Query;
    }
}
