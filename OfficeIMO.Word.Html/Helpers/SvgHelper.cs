namespace OfficeIMO.Word.Html {
    internal static class SvgHelper {
        internal static void AddSvg(WordParagraph paragraph, string svgContent, double? width, double? height, string description) {
            using var ms = new MemoryStream(Encoding.UTF8.GetBytes(svgContent));
            paragraph.AddImage(ms, "image.svg", (width is { } imageLength1 ? OfficeIMO.OfficeLength.Pixels(imageLength1) : (OfficeIMO.OfficeLength? )null), (height is { } imageLength2 ? OfficeIMO.OfficeLength.Pixels(imageLength2) : (OfficeIMO.OfficeLength? )null), options: new OfficeIMO.Word.WordImageOptions { Description = description });
        }
    }
}
