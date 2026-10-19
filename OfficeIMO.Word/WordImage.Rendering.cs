using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Word;

public partial class WordImage {
    // Rendering prefers the vector resource, but raw extraction retains the ordinary blip carrier.
    internal IEnumerable<byte[]> GetRenderingImageCandidates(CancellationToken cancellationToken) {
        const string relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var blip = GetBlip();
        if (blip != null) {
            OpenXmlPart owner = GetContainingPart();
            foreach (OpenXmlElement extension in blip.Descendants().Where(element => element.LocalName == "svgBlip" &&
                (element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2016/SVG/main" ||
                 element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main"))) {
                cancellationToken.ThrowIfCancellationRequested();
                string? id = extension.GetAttribute("embed", relationships).Value;
                if (id == null || id.Length == 0 || !owner.TryGetPartById(id, out OpenXmlPart? part) || part is not ImagePart svg) continue;
                using var stream = svg.GetStream(FileMode.Open, FileAccess.Read);
                yield return OfficeStreamReader.ReadAllBytes(stream, cancellationToken);
            }
        }
        using var fallback = OpenRead();
        yield return OfficeStreamReader.ReadAllBytes(fallback, cancellationToken);
    }
}
