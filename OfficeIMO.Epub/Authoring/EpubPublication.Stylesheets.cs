using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private bool ValidateStylesheetClosure(string owner, IReadOnlyDictionary<string, byte[]> entries,
        EpubManifestItem[] manifest, CancellationToken token) {
        var pending = new Queue<string>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        pending.Enqueue(owner);
        bool hasRemote = false;
        HtmlResourceManifest empty = HtmlResourcePipeline.BuildManifest(string.Empty);
        while (pending.Count != 0) {
            token.ThrowIfCancellationRequested();
            string path = pending.Dequeue();
            if (!visited.Add(path)) continue;
            if (_encryption.Any(encryption => encryption.Path == path))
                throw new NotSupportedException("Rewriting content that depends on protected stylesheets requires decryption.");
            if (!HtmlResourcePipeline.TryDecodeStylesheet(entries[path], "text/css", out string css))
                throw new InvalidDataException("A linked stylesheet could not be decoded: " + path);
            var document = new XDocument(new XElement(Html + "html", new XElement(Html + "head", new XElement(Html + "style", css)), new XElement(Html + "body")));
            hasRemote |= ValidateAuthoredResources(document, path, empty, manifest, token);
            HtmlExternalStylesheetAnalysis analysis = HtmlResourcePipeline.AnalyzeExternalStylesheet(css,
                new Uri("epub://package/" + EncodePath(path)), new HtmlResourcePipelineOptions(), includeInactiveResources: true);
            foreach (var import in analysis.Imports) {
                EpubReference reference = EpubReference.Resolve(path, import.Reference.Source);
                if (reference.Kind == EpubReferenceKind.Container) pending.Enqueue(reference.ContainerPath!);
            }
        }
        return hasRemote;
    }
}
