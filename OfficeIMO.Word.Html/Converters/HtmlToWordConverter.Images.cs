using AngleSharp.Html.Dom;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessImage(IHtmlImageElement img, WordDocument doc) {
            var src = img.Source;
            if (string.IsNullOrEmpty(src)) return;

            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;

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
                    doc.AddParagraph().AddImageFromBase64(base64, "image." + ext, width, height);
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                doc.AddParagraph().AddImage(uri.LocalPath, width, height);
            } else if (File.Exists(src)) {
                doc.AddParagraph().AddImage(src, width, height);
            } else {
                doc.AddImageFromUrl(src, width, height);
            }
        }
    }
}