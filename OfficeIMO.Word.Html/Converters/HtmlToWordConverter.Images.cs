using AngleSharp.Html.Dom;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessImage(IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options) {
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
                    doc.AddParagraph().AddImageFromBase64(base64, "image." + ext, width, height, description: alt);
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                doc.AddParagraph().AddImage(uri.LocalPath, width, height, description: alt);
            } else if (File.Exists(src)) {
                doc.AddParagraph().AddImage(src, width, height, description: alt);
            } else {
                var image = doc.AddImageFromUrl(src, width, height);
                image.Description = alt;
            }
        }
    }
}