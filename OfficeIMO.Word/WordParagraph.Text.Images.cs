using DocumentFormat.OpenXml;
using System.Text;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // The visible-text walker supplies offsets, including the placeholder occupied by a
    // page/column break. Images sharing an offset retain their source order.
    internal IReadOnlyList<(int Offset, WordImage Image)> GetPositionedImages() {
        var result = new List<(int Offset, WordImage Image)>();
        if (_run == null) return result;
        var images = new Dictionary<OpenXmlElement, WordImage>();
        foreach (WordImage image in EnumerateImages()) {
            OpenXmlElement? element = (OpenXmlElement?)image._Image ?? image._vmlShape;
            if (element != null) images[element] = image;
        }
        if (images.Count == 0) return result;
        AppendVisibleText(new StringBuilder(), _run, observeElement: (element, offset) => {
            if (images.TryGetValue(element, out WordImage? image)) result.Add((offset, image));
        });
        return result;
    }
}
