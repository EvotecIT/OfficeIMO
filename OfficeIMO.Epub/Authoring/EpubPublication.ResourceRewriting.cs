using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private byte[] RewritePublicationResource(string type, byte[] payload, string path, string destination,
        string oldPath, string newPath, CancellationToken token, ContentReferenceMap? map = null) {
        byte[] rewritten = payload;
        if (HasMediaType(type, "text/css")) {
            if (!HtmlResourcePipeline.TryDecodeStylesheet(payload, "text/css", out string css)) throw new InvalidDataException("Stylesheet cannot be decoded: " + path);
            string result = RewriteMovedCss(css, value => RewriteMovedReference(path, null, destination, null, value, oldPath, newPath, map), includeFragmentReferences: map != null);
            if (result != css) {
                // The shared decoder has consumed any source encoding. Emit matching UTF-8 bytes and declaration.
                if (result.TrimStart().StartsWith("@charset", StringComparison.OrdinalIgnoreCase)) {
                    int end = result.IndexOf(';');
                    if (end >= 0) result = "@charset \"UTF-8\";" + result.Substring(end + 1);
                }
                rewritten = new UTF8Encoding(false, true).GetBytes(result);
            }
        } else if (IsRenameXml(type)) {
            XDocument document = ParseXml(payload, _maximumEntryBytes);
            XName expected = HasMediaType(type, "application/xhtml+xml") ? Html + "html" : HasMediaType(type, "image/svg+xml") ?
                XName.Get("svg", "http://www.w3.org/2000/svg") : HasMediaType(type, "application/smil+xml") ?
                XName.Get("smil", "http://www.w3.org/ns/SMIL") : Ncx + "ncx";
            if (document.Root?.Name != expected) throw new InvalidDataException("Resource XML root does not match its media type: " + path);
            if (RewriteMovedXml(document, path, destination, oldPath, newPath, token, map))
                rewritten = SerializeXml(document, _maximumEntryBytes);
        } else if (!IsRenameLeaf(type)) {
            throw new NotSupportedException("Resource renaming cannot inspect references in media type: " + type);
        }
        return rewritten;
    }
}
