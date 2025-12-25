using AngleSharp.Html.Dom;
using System.Globalization;
using System.Net.Http;
using System.Threading;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word.Html {
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
            width ??= TryParsePixelValue(img.GetAttribute("width"));
            height ??= TryParsePixelValue(img.GetAttribute("height"));
            var alt = img.AlternativeText ?? string.Empty;

            if (options.ImageProcessing == ImageProcessingMode.EmbedDataUriOnly && !src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                InsertAltText(currentParagraph, headerFooter, doc, alt);
                return;
            }

            WordParagraph? paragraph = currentParagraph;

            if (horizontalAlignment == null && _imageCache.TryGetValue(src, out var cached)) {
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                cached.Clone(paragraph);
                return;
            }

            WordImage image;
            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                if (!TryHandleDataImage(src, doc, ref paragraph, headerFooter, width, height, wrap, alt, out image)) {
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else if (options.ImageProcessing == ImageProcessingMode.LinkExternal) {
                if (!TryHandleExternalImage(src, doc, ref paragraph, headerFooter, width, height, wrap, alt, out image)) {
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                try {
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(uri.LocalPath, width, height, wrap, description: alt);
                    image = paragraph.Image!;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from file '{uri.LocalPath}': {ex.Message}");
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else if (File.Exists(src)) {
                try {
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(src, width, height, wrap, description: alt);
                    image = paragraph.Image!;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from file '{src}': {ex.Message}");
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else {
                try {
                    var data = FetchBytes(new Uri(src));
                    using var ms = new MemoryStream(data);
                    string fileName = GetFileNameFromUri(src);
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(ms, fileName, width, height, wrap, description: alt);
                    image = paragraph.Image!;
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load image from '{src}': {ex.Message}");
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
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
            width ??= TryParsePixelValue(img.GetAttribute("width"));
            height ??= TryParsePixelValue(img.GetAttribute("height"));
            var alt = img.AlternativeText;

            if (options.ImageProcessing == ImageProcessingMode.EmbedDataUriOnly && !src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
                return;
            }

            WordParagraph paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());

            string svgContent;
            if (src.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                if (!TryGetDataUriContent(src, out var meta, out var data, out var isBase64)) {
                    return;
                }
                if (isBase64) {
                    var bytes = System.Convert.FromBase64String(data);
                    svgContent = Encoding.UTF8.GetString(bytes);
                } else {
                    svgContent = Uri.UnescapeDataString(data);
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                svgContent = File.ReadAllText(uri.LocalPath);
            } else if (File.Exists(src)) {
                svgContent = File.ReadAllText(src);
            } else {
                svgContent = FetchString(new Uri(src));
            }

            try {
                SvgHelper.AddSvg(paragraph, svgContent, width, height, alt ?? string.Empty);
                _imageCache[src] = paragraph.Image!;
            } catch (Exception ex) {
                Console.WriteLine($"Failed to embed SVG: {ex.Message}");
                if (!string.IsNullOrEmpty(alt)) {
                    paragraph.AddText(alt ?? string.Empty);
                }
            }
        }

        private static double? TryParsePixelValue(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }
            var trimmed = value!.Trim().ToLowerInvariant();
            if (trimmed.EndsWith("px", StringComparison.Ordinal)) {
                trimmed = trimmed.Substring(0, trimmed.Length - 2);
            }
            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var result)) {
                return result > 0 ? result : null;
            }
            return null;
        }

        private static string GetFileNameFromUri(string src) {
            try {
                var uriSrc = new Uri(src);
                var fileName = Path.GetFileName(uriSrc.LocalPath);
                return string.IsNullOrEmpty(fileName) ? "image" : fileName;
            } catch (UriFormatException) {
                return "image";
            }
        }

        private static void InsertAltText(WordParagraph? currentParagraph, WordHeaderFooter? headerFooter, WordDocument doc, string alt) {
            if (string.IsNullOrEmpty(alt)) {
                return;
            }
            var paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
            paragraph.AddText(alt);
        }

        private bool TryHandleExternalImage(string src, WordDocument doc, ref WordParagraph? paragraph, WordHeaderFooter? headerFooter, double? width, double? height, WrapTextImage wrap, string alt, out WordImage image) {
            image = null!;
            if (!width.HasValue || !height.HasValue) {
                return false;
            }

            if (Uri.TryCreate(src, UriKind.Absolute, out var uri)) {
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                paragraph.AddImage(uri, width.Value, height.Value, wrap, description: alt);
                image = paragraph.Image!;
                return true;
            }

            if (File.Exists(src)) {
                var fileUri = new Uri(Path.GetFullPath(src));
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                paragraph.AddImage(fileUri, width.Value, height.Value, wrap, description: alt);
                image = paragraph.Image!;
                return true;
            }

            return false;
        }

        private bool TryHandleDataImage(string src, WordDocument doc, ref WordParagraph? paragraph, WordHeaderFooter? headerFooter, double? width, double? height, WrapTextImage wrap, string alt, out WordImage image) {
            image = null!;
            if (!TryGetDataUriContent(src, out var meta, out var data, out var isBase64)) {
                return false;
            }
            try {
                var ext = "png";
                var parts = meta.Split(new[] { ';', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2) {
                    ext = parts[1];
                }
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                if (isBase64) {
                    paragraph.AddImageFromBase64(data, "image." + ext, width, height, wrap, description: alt);
                } else {
                    if (!meta.Contains("svg+xml", StringComparison.OrdinalIgnoreCase)) {
                        return false;
                    }
                    var svgContent = Uri.UnescapeDataString(data);
                    SvgHelper.AddSvg(paragraph, svgContent, width, height, alt);
                }
                image = paragraph.Image!;
                return true;
            } catch (Exception ex) {
                Console.WriteLine($"Failed to decode data-image: {ex.Message}");
                return false;
            }
        }

        private static bool TryGetDataUriContent(string src, out string meta, out string data, out bool isBase64) {
            meta = string.Empty;
            data = string.Empty;
            isBase64 = false;
            var commaIndex = src.IndexOf(',');
            if (commaIndex <= 0) {
                return false;
            }
            meta = src.Substring(5, commaIndex - 5);
            data = src.Substring(commaIndex + 1);
            isBase64 = meta.IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0;
            return true;
        }

        private byte[] FetchBytes(Uri uri) {
            using var cts = _resourceTimeout.HasValue
                ? CancellationTokenSource.CreateLinkedTokenSource(_cancellationToken)
                : null;
            var token = cts?.Token ?? _cancellationToken;
            if (cts != null && _resourceTimeout.HasValue) {
                cts.CancelAfter(_resourceTimeout.Value);
            }
            using var request = new HttpRequestMessage(HttpMethod.Get, uri);
            using var response = _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();
            return response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
        }

        private string FetchString(Uri uri) {
            using var cts = _resourceTimeout.HasValue
                ? CancellationTokenSource.CreateLinkedTokenSource(_cancellationToken)
                : null;
            var token = cts?.Token ?? _cancellationToken;
            if (cts != null && _resourceTimeout.HasValue) {
                cts.CancelAfter(_resourceTimeout.Value);
            }
            using var request = new HttpRequestMessage(HttpMethod.Get, uri);
            using var response = _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();
            return response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        }

        private void ProcessSvgElement(AngleSharp.Dom.IElement svg, WordDocument doc, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            double? width = null;
            double? height = null;
            if (double.TryParse(svg.GetAttribute("width")?.Replace("px", string.Empty), out var w)) width = w;
            if (double.TryParse(svg.GetAttribute("height")?.Replace("px", string.Empty), out var h)) height = h;

            var paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : section.AddParagraph());
            try {
                SvgHelper.AddSvg(paragraph, svg.OuterHtml, width, height, string.Empty);
            } catch (Exception ex) {
                Console.WriteLine($"Failed to embed inline SVG: {ex.Message}");
            }
        }
    }
}
