using DocumentFormat.OpenXml;
using System.Text;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // The visible-text walker supplies offsets, including the placeholder occupied by a
    // page/column break. Images sharing an offset retain their source order.
    internal IReadOnlyList<(int Offset, WordImage Image)> GetPositionedImages() {
        var result = new List<(int Offset, WordImage Image)>();
        OpenXmlElement? contentRun = _visibleRun ?? _run;
        if (contentRun == null) return result;
        var images = new Dictionary<OpenXmlElement, WordImage>();
        foreach (WordImage image in EnumerateImages()) {
            OpenXmlElement? element = (OpenXmlElement?)image._Image ?? image._vmlShape;
            if (element != null) images[element] = image;
        }
        if (images.Count == 0) return result;
        if (_visibleRun != null && _visibleRunSourceChildren != null) {
            var projectedImages = new Dictionary<OpenXmlElement, WordImage>();
            for (int index = 0; index < _visibleRunSourceChildren.Count; index++) {
                OpenXmlElement sourceChild = _visibleRunSourceChildren[index];
                OpenXmlElement projectedChild = _visibleRun.ChildElements[index];
                OpenXmlElement[] sourceNodes = sourceChild.Descendants().Prepend(sourceChild).ToArray();
                OpenXmlElement[] projectedNodes = projectedChild.Descendants().Prepend(projectedChild).ToArray();
                for (int node = 0; node < sourceNodes.Length && node < projectedNodes.Length; node++) {
                    if (images.TryGetValue(sourceNodes[node], out WordImage? image))
                        projectedImages[projectedNodes[node]] = image;
                }
            }
            images = projectedImages;
        }
        AppendVisibleText(new StringBuilder(), contentRun, observeElement: (element, offset) => {
            if (images.TryGetValue(element, out WordImage? image)) result.Add((offset, image));
        });
        return result;
    }
}
