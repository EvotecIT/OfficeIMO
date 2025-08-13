using OfficeIMO.Word;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Html.Helpers {
    internal static class SvgHelper {
        internal static void AddSvg(WordParagraph paragraph, string svgContent, double? width, double? height, string description) {
            using var ms = new MemoryStream(Encoding.UTF8.GetBytes(svgContent));
            paragraph.AddImage(ms, "image.svg", width, height, description: description);
        }
    }
}
