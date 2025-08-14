using AngleSharp.Html.Dom;
using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Helpers;
using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private async Task ProcessImageAsync(IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter, CancellationToken cancellationToken) {
            var src = img.GetAttribute("src");
            if (string.IsNullOrEmpty(src)) return;

            if (!src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase) && !Uri.TryCreate(src, UriKind.Absolute, out _)) {
                if (!string.IsNullOrEmpty(options.BasePath)) {
                    src = Path.Combine(options.BasePath, src);
                } else if (img.BaseUrl != null && Uri.TryCreate(img.BaseUrl.Href, UriKind.Absolute, out var baseUri) && !string.Equals(img.BaseUrl.Href, "http://localhost/", StringComparison.OrdinalIgnoreCase)) {
                    src = new Uri(baseUri, src).ToString();
                }
            }

            if (src.EndsWith(".svg", StringComparison.OrdinalIgnoreCase) || src.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                await ProcessSvgImageAsync(src, img, doc, options, currentParagraph, headerFooter, cancellationToken).ConfigureAwait(false);
                return;
            }

            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;
            var alt = img.AlternativeText;

            WordParagraph? paragraph = currentParagraph;

            if (_imageCache.TryGetValue(src, out var cached)) {
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                var drawingField = typeof(WordImage).GetField("_Image", BindingFlags.Instance | BindingFlags.NonPublic);
                var drawing = (DocumentFormat.OpenXml.Wordprocessing.Drawing)drawingField!.GetValue(cached);
                var clone = (DocumentFormat.OpenXml.Wordprocessing.Drawing)drawing.CloneNode(true);
                var run = new Run(clone);
                var paragraphField = typeof(WordParagraph).GetField("_paragraph", BindingFlags.Instance | BindingFlags.NonPublic);
                var p = (Paragraph)paragraphField!.GetValue(paragraph);
                p.Append(run);
                return;
            }

            WordImage image;
            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                var commaIndex = src.IndexOf(',');
                if (commaIndex > 0) {
                    var meta = src.Substring(5, commaIndex - 5); // e.g., image/png;base64
                    var base64 = src.Substring(commaIndex + 1);
                    var ext = "png";
                    var parts = meta.Split(new[] { ';', '/' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2) {
                        ext = parts[1];
                    }
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImageFromBase64(base64, "image." + ext, width, height, description: alt);
                    image = paragraph.Image;
                } else {
                    return;
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                paragraph.AddImage(uri.LocalPath, width, height, description: alt);
                image = paragraph.Image;
            } else if (File.Exists(src)) {
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                paragraph.AddImage(src, width, height, description: alt);
                image = paragraph.Image;
            } else {
                try {
                    var data = await _imageDownloader.DownloadAsync(src, cancellationToken).ConfigureAwait(false);
                    if (data == null) {
                        throw new Exception("Download failed");
                    }
                    using var ms = new MemoryStream(data);
                    string fileName = "image";
                    try {
                        var uriSrc = new Uri(src);
                        fileName = Path.GetFileName(uriSrc.LocalPath);
                        if (string.IsNullOrEmpty(fileName)) fileName = "image";
                    } catch (UriFormatException) {
                        // ignore
                    }
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(ms, fileName, width, height, description: alt);
                    image = paragraph.Image;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from '{src}': {ex.Message}");
                    if (!string.IsNullOrEmpty(alt)) {
                        paragraph ??= currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
                        paragraph.AddText(alt);
                    }
                    return;
                }
            }

            _imageCache[src] = image;
        }

        private async Task ProcessSvgImageAsync(string src, IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter, CancellationToken cancellationToken) {
            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;
            var alt = img.AlternativeText;

            WordParagraph paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());

            string svgContent;
            if (src.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                var commaIndex = src.IndexOf(',');
                if (commaIndex < 0) return;
                var base64 = src.Substring(commaIndex + 1);
                var bytes = Convert.FromBase64String(base64);
                svgContent = Encoding.UTF8.GetString(bytes);
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                svgContent = File.ReadAllText(uri.LocalPath);
            } else if (File.Exists(src)) {
                svgContent = File.ReadAllText(src);
            } else {
                var data = await _imageDownloader.DownloadAsync(src, cancellationToken).ConfigureAwait(false);
                svgContent = data != null ? Encoding.UTF8.GetString(data) : string.Empty;
            }

            SvgHelper.AddSvg(paragraph, svgContent, width, height, alt);
            _imageCache[src] = paragraph.Image;
        }

        private void ProcessSvgElement(AngleSharp.Dom.IElement svg, WordDocument doc, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            double? width = null;
            double? height = null;
            if (double.TryParse(svg.GetAttribute("width")?.Replace("px", string.Empty), out var w)) width = w;
            if (double.TryParse(svg.GetAttribute("height")?.Replace("px", string.Empty), out var h)) height = h;

            var paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : section.AddParagraph());
            SvgHelper.AddSvg(paragraph, svg.OuterHtml, width, height, string.Empty);
        }
    }
}