using AngleSharp.Html.Dom;
using AngleSharp.Dom;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Helpers;
using System;
using System.IO;
using System.Net.Http;
using System.Text;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessImage(IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            var src = img.GetAttribute("src") ?? string.Empty;
            if (string.IsNullOrEmpty(src)) return;

            var decl = _inlineParser.ParseDeclaration(img.GetAttribute("style") ?? string.Empty);
            var floatVal = decl.GetPropertyValue("float")?.Trim().ToLowerInvariant();
            var wrap = WrapTextImage.InLineWithText;
            string? horizontalAlignment = null;
            if (floatVal == "left") {
                wrap = WrapTextImage.Square;
                horizontalAlignment = "left";
            } else if (floatVal == "right") {
                wrap = WrapTextImage.Square;
                horizontalAlignment = "right";
            }

            if (!src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase) && !Uri.TryCreate(src, UriKind.Absolute, out _)) {
                if (!string.IsNullOrEmpty(options.BasePath)) {
                    src = Path.Combine(options.BasePath, src);
                } else if (img.BaseUrl != null && Uri.TryCreate(img.BaseUrl.Href, UriKind.Absolute, out var baseUri) && !string.Equals(img.BaseUrl.Href, "http://localhost/", StringComparison.OrdinalIgnoreCase)) {
                    src = new Uri(baseUri, src).ToString();
                }
            }

            if (src.EndsWith(".svg", StringComparison.OrdinalIgnoreCase) || src.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                ProcessSvgImage(src, img, doc, options, currentParagraph, headerFooter);
                return;
            }

            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;
            var alt = img.AlternativeText ?? string.Empty;

            WordParagraph? paragraph = currentParagraph;

            if (horizontalAlignment == null && _imageCache.TryGetValue(src, out var cached)) {
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                cached.Clone(paragraph);
                return;
            }

            WordImage image;
            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                var commaIndex = src.IndexOf(',');
                if (commaIndex > 0) {
                    try {
                        var meta = src.Substring(5, commaIndex - 5); // e.g., image/png;base64
                        var base64 = src.Substring(commaIndex + 1);
                        var ext = "png";
                        var parts = meta.Split(new[] { ';', '/' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 2) {
                            ext = parts[1];
                        }
                        paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                        paragraph.AddImageFromBase64(base64, "image." + ext, width, height, wrap, description: alt);
                        image = paragraph.Image!;
                    } catch (Exception ex) {
                        Console.WriteLine($"Failed to decode data-image: {ex.Message}");
                        if (!string.IsNullOrEmpty(alt)) {
                            paragraph ??= currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
                            paragraph.AddText(alt);
                        }
                        return;
                    }
                } else {
                    return;
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                try {
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(uri.LocalPath, width, height, wrap, description: alt);
                    image = paragraph.Image!;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from file '{uri.LocalPath}': {ex.Message}");
                    if (!string.IsNullOrEmpty(alt)) {
                        paragraph ??= currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
                        paragraph.AddText(alt);
                    }
                    return;
                }
            } else if (File.Exists(src)) {
                try {
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(src, width, height, wrap, description: alt);
                    image = paragraph.Image!;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from file '{src}': {ex.Message}");
                    if (!string.IsNullOrEmpty(alt)) {
                        paragraph ??= currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
                        paragraph.AddText(alt);
                    }
                    return;
                }
            } else {
                try {
                    using HttpClient client = new HttpClient();
                    var data = client.GetByteArrayAsync(src).GetAwaiter().GetResult();
                    using var ms = new MemoryStream(data);
                    string fileName = "image";
                    try {
                        var uriSrc = new Uri(src);
                        fileName = Path.GetFileName(uriSrc.LocalPath) ?? "image";
                        if (string.IsNullOrEmpty(fileName)) fileName = "image";
                    } catch (UriFormatException) {
                        // ignore
                    }
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(ms, fileName, width, height, wrap, description: alt);
                    image = paragraph.Image!;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from '{src}': {ex.Message}");
                    if (!string.IsNullOrEmpty(alt)) {
                        paragraph ??= currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
                        paragraph.AddText(alt);
                    }
                    return;
                }
            }

            if (horizontalAlignment != null) {
                var hPos = image.horizontalPosition;
                hPos?.GetFirstChild<Wp.PositionOffset>()?.Remove();
                if (hPos != null) {
                    hPos.HorizontalAlignment = new Wp.HorizontalAlignment() { Text = horizontalAlignment };
                }
            }

            if (horizontalAlignment == null) {
                _imageCache[src] = image;
            }
        }

        private void ProcessSvgImage(string src, IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;
            var alt = img.AlternativeText;

            WordParagraph paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());

            string svgContent;
            if (src.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                var commaIndex = src.IndexOf(',');
                if (commaIndex < 0) return;
                var base64 = src.Substring(commaIndex + 1);
                var bytes = System.Convert.FromBase64String(base64);
                svgContent = Encoding.UTF8.GetString(bytes);
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                svgContent = File.ReadAllText(uri.LocalPath);
            } else if (File.Exists(src)) {
                svgContent = File.ReadAllText(src);
            } else {
                using HttpClient client = new HttpClient();
                svgContent = client.GetStringAsync(src).GetAwaiter().GetResult();
            }

            try {
                SvgHelper.AddSvg(paragraph, svgContent, width, height, alt ?? string.Empty);
                _imageCache[src] = paragraph.Image!;
            } catch (System.Exception ex) {
                System.Console.WriteLine($"Failed to embed SVG: {ex.Message}");
                if (!string.IsNullOrEmpty(alt)) {
                    paragraph.AddText(alt ?? string.Empty);
                }
            }
        }

        private void ProcessSvgElement(AngleSharp.Dom.IElement svg, WordDocument doc, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            double? width = null;
            double? height = null;
            if (double.TryParse(svg.GetAttribute("width")?.Replace("px", string.Empty), out var w)) width = w;
            if (double.TryParse(svg.GetAttribute("height")?.Replace("px", string.Empty), out var h)) height = h;

            var paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : section.AddParagraph());
            try {
                SvgHelper.AddSvg(paragraph, svg.OuterHtml, width, height, string.Empty);
            } catch (System.Exception ex) {
                System.Console.WriteLine($"Failed to embed inline SVG: {ex.Message}");
            }
        }
    }
}
