using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Net.Http;
using System.Reflection;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessImage(IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph) {
            var src = img.GetAttribute("src");
            if (string.IsNullOrEmpty(src)) return;

            if (!src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase) && !Uri.TryCreate(src, UriKind.Absolute, out _)) {
                if (!string.IsNullOrEmpty(options.BasePath)) {
                    src = Path.Combine(options.BasePath, src);
                } else if (img.BaseUrl != null && Uri.TryCreate(img.BaseUrl.Href, UriKind.Absolute, out var baseUri) && !string.Equals(img.BaseUrl.Href, "http://localhost/", StringComparison.OrdinalIgnoreCase)) {
                    src = new Uri(baseUri, src).ToString();
                }
            }

            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;
            var alt = img.AlternativeText;

            WordParagraph? paragraph = currentParagraph;

            if (_imageCache.TryGetValue(src, out var cached)) {
                paragraph ??= doc.AddParagraph();
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
                    paragraph ??= doc.AddParagraph();
                    paragraph.AddImageFromBase64(base64, "image." + ext, width, height, description: alt);
                    image = paragraph.Image;
                } else {
                    return;
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                paragraph ??= doc.AddParagraph();
                paragraph.AddImage(uri.LocalPath, width, height, description: alt);
                image = paragraph.Image;
            } else if (File.Exists(src)) {
                paragraph ??= doc.AddParagraph();
                paragraph.AddImage(src, width, height, description: alt);
                image = paragraph.Image;
            } else {
                try {
                    using HttpClient client = new HttpClient();
                    var data = client.GetByteArrayAsync(src).GetAwaiter().GetResult();
                    using var ms = new MemoryStream(data);
                    string fileName = "image";
                    try {
                        var uriSrc = new Uri(src);
                        fileName = Path.GetFileName(uriSrc.LocalPath);
                        if (string.IsNullOrEmpty(fileName)) fileName = "image";
                    } catch (UriFormatException) {
                        // ignore
                    }
                    paragraph ??= doc.AddParagraph();
                    paragraph.AddImage(ms, fileName, width, height, description: alt);
                    image = paragraph.Image;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from '{src}': {ex.Message}");
                    if (!string.IsNullOrEmpty(alt)) {
                        paragraph ??= currentParagraph ?? doc.AddParagraph();
                        paragraph.AddText(alt);
                    }
                    return;
                }
            }

            _imageCache[src] = image;
        }
    }
}