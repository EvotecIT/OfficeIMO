using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using System.Globalization;
using System.Net.Http;
using System.Threading;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static readonly HtmlUrlPolicy ImageSourceResolutionPolicy = CreateImageSourceResolutionPolicy();

        private void ProcessImage(IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            var src = ResolveWordImageSource(img, options);
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

            var alt = img.AlternativeText ?? string.Empty;
            var title = img.GetAttribute("title") ?? string.Empty;
            if (options.ImageProcessing == ImageProcessingMode.EmbedDataUriOnly && !src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                AddDiagnostic(options, "ImageSkippedByPolicy", "External image was skipped because only data URI images are enabled.", src);
                InsertAltText(currentParagraph, headerFooter, doc, alt);
                return;
            }

            if (!TryApplyImageSourcePolicy(src, options, currentParagraph, headerFooter, doc, alt)) {
                return;
            }

            if (src.EndsWith(".svg", StringComparison.OrdinalIgnoreCase) || src.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                ProcessSvgImage(src, img, doc, options, currentParagraph, headerFooter);
                return;
            }

            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;
            width ??= TryResolveImagePercentWidth(decl.GetPropertyValue("width"), doc);
            width ??= TryResolveImagePercentWidth(img.GetAttribute("width"), doc);
            width ??= TryParsePixelValue(img.GetAttribute("width"));
            height ??= TryParsePixelValue(img.GetAttribute("height"));

            WordParagraph? paragraph = currentParagraph;

            if (horizontalAlignment == null && _imageCache.TryGetValue(src, out var cached)) {
                paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                var clonedImage = cached.Clone(paragraph);
                ApplyImageMetadata(clonedImage, alt, title);
                return;
            }

            WordImage image;
            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                if (!TryHandleDataImage(src, doc, options, ref paragraph, headerFooter, width, height, wrap, alt, out image)) {
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else if (options.ImageProcessing == ImageProcessingMode.LinkExternal) {
                if (!TryHandleExternalImage(src, doc, ref paragraph, headerFooter, width, height, wrap, alt, out image)) {
                    AddDiagnostic(options, "ImageLinkSkipped", "External image link could not be created. LinkExternal requires a resolvable source with explicit width and height.", src);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                long reservedBytes = 0;
                try {
                    reservedBytes = EnsureFileWithinImageLimits(uri.LocalPath, options);
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(uri.LocalPath, width, height, wrap, description: alt);
                    reservedBytes = 0;
                    image = paragraph.Image!;
                } catch (HtmlResourceLimitException ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageResourceTooLarge", "Image file exceeded the configured byte limit and was replaced with alt text when available.", uri.LocalPath, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                } catch (HtmlResourceTotalLimitException ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageResourceBudgetExceeded", "Image file exceeded the configured total byte budget and was replaced with alt text when available.", uri.LocalPath, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                } catch (Exception ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageLoadFailed", "Image file could not be loaded and was replaced with alt text when available.", uri.LocalPath, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else if (File.Exists(src)) {
                long reservedBytes = 0;
                try {
                    reservedBytes = EnsureFileWithinImageLimits(src, options);
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(src, width, height, wrap, description: alt);
                    reservedBytes = 0;
                    image = paragraph.Image!;
                } catch (HtmlResourceLimitException ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageResourceTooLarge", "Image file exceeded the configured byte limit and was replaced with alt text when available.", src, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                } catch (HtmlResourceTotalLimitException ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageResourceBudgetExceeded", "Image file exceeded the configured total byte budget and was replaced with alt text when available.", src, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                } catch (Exception ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageLoadFailed", "Image file could not be loaded and was replaced with alt text when available.", src, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                }
            } else {
                long reservedBytes = 0;
                try {
                    var data = FetchBytes(new Uri(src), options);
                    reservedBytes = data.LongLength;
                    using var ms = new MemoryStream(data);
                    string fileName = GetFileNameFromUri(src);
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImage(ms, fileName, width, height, wrap, description: alt);
                    reservedBytes = 0;
                    image = paragraph.Image!;
                } catch (HtmlResourceLimitException ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageResourceTooLarge", "Image resource exceeded the configured byte limit and was replaced with alt text when available.", src, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                } catch (HtmlResourceTotalLimitException ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageResourceBudgetExceeded", "Image resource exceeded the configured total byte budget and was replaced with alt text when available.", src, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                } catch (HtmlResourceContentTypeException ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageContentTypeRejected", "Image resource content type is not allowed and was replaced with alt text when available.", src, ex);
                    InsertAltText(currentParagraph, headerFooter, doc, alt);
                    return;
                } catch (Exception ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    AddDiagnostic(options, "ImageLoadFailed", "Image resource could not be loaded and was replaced with alt text when available.", src, ex);
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

            ApplyImageMetadata(image, alt, title);

            if (horizontalAlignment == null) {
                _imageCache[src] = image;
            }
        }

        private void ProcessSvgImage(string src, IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            var decl = _inlineParser.ParseDeclaration(img.GetAttribute("style") ?? string.Empty);
            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;
            width ??= TryResolveImagePercentWidth(decl.GetPropertyValue("width"), doc);
            width ??= TryResolveImagePercentWidth(img.GetAttribute("width"), doc);
            width ??= TryParsePixelValue(img.GetAttribute("width"));
            height ??= TryParsePixelValue(img.GetAttribute("height"));
            var alt = img.AlternativeText;
            var title = img.GetAttribute("title") ?? string.Empty;

            if (options.ImageProcessing == ImageProcessingMode.EmbedDataUriOnly && !src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                AddDiagnostic(options, "ImageSkippedByPolicy", "External SVG image was skipped because only data URI images are enabled.", src);
                InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
                return;
            }

            long reservedBytes = 0;
            try {
                string svgContent;
                if (src.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                    if (!HtmlImageDataUri.TryParse(src, out var dataUri)) {
                        AddDiagnostic(options, "SvgDataUriInvalid", "SVG data URI could not be parsed and was skipped.", src);
                        InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
                        return;
                    }
                    EnsureImageContentTypeAllowed(dataUri.MediaType, options);
                    if (dataUri.IsBase64) {
                        var estimatedBytes = dataUri.EstimateDecodedByteCount();
                        if (options.MaxImageBytes.HasValue && estimatedBytes > options.MaxImageBytes.Value) {
                            AddDiagnostic(options, "ImageResourceTooLarge", "SVG data URI exceeded the configured byte limit and was replaced with alt text when available.", "data:image/svg+xml");
                            InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
                            return;
                        }
                        if (!TryReserveImageBytes(estimatedBytes, options, "data:image/svg+xml")) {
                            InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
                            return;
                        }
                        reservedBytes = estimatedBytes;
                        var bytes = dataUri.DecodeBytes();
                        svgContent = Encoding.UTF8.GetString(bytes);
                    } else {
                        svgContent = dataUri.DecodeText();
                        var svgByteCount = Encoding.UTF8.GetByteCount(svgContent);
                        if (options.MaxImageBytes.HasValue && svgByteCount > options.MaxImageBytes.Value) {
                            AddDiagnostic(options, "ImageResourceTooLarge", "SVG data URI exceeded the configured byte limit and was replaced with alt text when available.", "data:image/svg+xml");
                            InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
                            return;
                        }
                        if (!TryReserveImageBytes(svgByteCount, options, "data:image/svg+xml")) {
                            InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
                            return;
                        }
                        reservedBytes = svgByteCount;
                    }
                } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                    reservedBytes = EnsureFileWithinImageLimits(uri.LocalPath, options);
                    svgContent = File.ReadAllText(uri.LocalPath);
                } else if (File.Exists(src)) {
                    reservedBytes = EnsureFileWithinImageLimits(src, options);
                    svgContent = File.ReadAllText(src);
                } else {
                    svgContent = FetchString(new Uri(src), options, out reservedBytes);
                }

                try {
                    var paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
                    SvgHelper.AddSvg(paragraph, svgContent, width, height, alt ?? string.Empty);
                    if (paragraph.Image != null) {
                        ApplyImageMetadata(paragraph.Image, alt ?? string.Empty, title);
                    }
                    _imageCache[src] = paragraph.Image!;
                    reservedBytes = 0;
                } catch (Exception ex) {
                    ReleaseImageBytes(reservedBytes, options);
                    reservedBytes = 0;
                    AddDiagnostic(options, "SvgEmbedFailed", "SVG image could not be embedded and was replaced with alt text when available.", src, ex);
                    if (!string.IsNullOrEmpty(alt)) {
                        var paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph());
                        paragraph.AddText(alt ?? string.Empty);
                    }
                }
            } catch (HtmlResourceLimitException ex) {
                ReleaseImageBytes(reservedBytes, options);
                AddDiagnostic(options, "ImageResourceTooLarge", "SVG image resource exceeded the configured byte limit and was replaced with alt text when available.", src, ex);
                InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
            } catch (HtmlResourceTotalLimitException ex) {
                ReleaseImageBytes(reservedBytes, options);
                AddDiagnostic(options, "ImageResourceBudgetExceeded", "SVG image resource exceeded the configured total byte budget and was replaced with alt text when available.", src, ex);
                InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
            } catch (HtmlResourceContentTypeException ex) {
                ReleaseImageBytes(reservedBytes, options);
                AddDiagnostic(options, "ImageContentTypeRejected", "SVG image resource content type is not allowed and was replaced with alt text when available.", src, ex);
                InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
            } catch (Exception ex) {
                ReleaseImageBytes(reservedBytes, options);
                AddDiagnostic(options, "SvgLoadFailed", "SVG image could not be loaded and was replaced with alt text when available.", src, ex);
                InsertAltText(currentParagraph, headerFooter, doc, alt ?? string.Empty);
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

        private static double? TryResolveImagePercentWidth(string? value, WordDocument doc) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            var trimmed = value!.Trim();
            if (!trimmed.EndsWith("%", StringComparison.Ordinal)) {
                return null;
            }

            if (!double.TryParse(trimmed.Substring(0, trimmed.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var percent) || percent <= 0) {
                return null;
            }

            var section = doc.Sections.Count > 0 ? doc.Sections[doc.Sections.Count - 1] : null;
            var pageWidthTwips = section?.PageSettings.Width?.Value ?? WordPageSizes.A4.Width!.Value;
            var leftMarginTwips = section?.Margins.Left?.Value ?? 1440U;
            var rightMarginTwips = section?.Margins.Right?.Value ?? 1440U;
            var contentWidthTwips = Math.Max(0D, pageWidthTwips - leftMarginTwips - rightMarginTwips);
            if (contentWidthTwips <= 0) {
                return null;
            }

            return contentWidthTwips / 15D * (percent / 100D);
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

        private static void ApplyImageMetadata(WordImage image, string alt, string? title) {
            image.Description = alt;
            image.Title = string.IsNullOrEmpty(title) ? null : title;
        }

        private long EnsureFileWithinImageLimits(string path, HtmlToWordOptions options) {
            var length = new FileInfo(path).Length;
            if (options.MaxImageBytes.HasValue && length > options.MaxImageBytes.Value) {
                throw new HtmlResourceLimitException($"Resource length {length} bytes exceeds limit {options.MaxImageBytes.Value} bytes.");
            }
            ReserveImageBytes(length, options);
            return length;
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

        private bool TryHandleDataImage(string src, WordDocument doc, HtmlToWordOptions options, ref WordParagraph? paragraph, WordHeaderFooter? headerFooter, double? width, double? height, WrapTextImage wrap, string alt, out WordImage image) {
            image = null!;
            if (!HtmlImageDataUri.TryParse(src, out var dataUri)) {
                AddDiagnostic(options, "ImageDataUriInvalid", "Image data URI could not be parsed and was skipped.", src);
                return false;
            }
            long reservedBytes = 0;
            try {
                if (!IsImageContentTypeAllowed(dataUri.MediaType, options)) {
                    AddDiagnostic(options, "ImageContentTypeRejected", "Image data URI content type is not allowed and was replaced with alt text when available.", "data:" + dataUri.MediaType);
                    return false;
                }

                string ext = dataUri.FileExtension.TrimStart('.');
                if (dataUri.IsBase64) {
                    var estimatedBytes = dataUri.EstimateDecodedByteCount();
                    if (options.MaxImageBytes.HasValue && estimatedBytes > options.MaxImageBytes.Value) {
                        AddDiagnostic(options, "ImageResourceTooLarge", "Image data URI exceeded the configured byte limit and was replaced with alt text when available.", "data:image");
                        return false;
                    }
                    if (!TryReserveImageBytes(estimatedBytes, options, "data:image")) {
                        return false;
                    }
                    reservedBytes = estimatedBytes;
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    paragraph.AddImageFromBase64(dataUri.Data, "image." + ext, width, height, wrap, description: alt);
                } else {
                    if (!dataUri.MediaType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                        AddDiagnostic(options, "ImageDataUriUnsupported", "Non-base64 data URI image was skipped because only SVG text data URIs are supported.", src);
                        return false;
                    }
                    var svgContent = dataUri.DecodeText();
                    if (options.MaxImageBytes.HasValue && Encoding.UTF8.GetByteCount(svgContent) > options.MaxImageBytes.Value) {
                        AddDiagnostic(options, "ImageResourceTooLarge", "SVG data URI exceeded the configured byte limit and was replaced with alt text when available.", "data:image/svg+xml");
                        return false;
                    }
                    var svgByteCount = Encoding.UTF8.GetByteCount(svgContent);
                    if (!TryReserveImageBytes(svgByteCount, options, "data:image/svg+xml")) {
                        return false;
                    }
                    reservedBytes = svgByteCount;
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    SvgHelper.AddSvg(paragraph, svgContent, width, height, alt);
                }
                image = paragraph.Image!;
                reservedBytes = 0;
                return true;
            } catch (Exception ex) {
                ReleaseImageBytes(reservedBytes, options);
                AddDiagnostic(options, "ImageDataUriInvalid", "Image data URI could not be decoded or embedded and was replaced with alt text when available.", src, ex);
                return false;
            }
        }

        private static Uri? ResolveImageBaseUri(IHtmlImageElement img, HtmlToWordOptions options) {
            if (!string.IsNullOrEmpty(options.BasePath)) {
                return null;
            }

            if (img.BaseUrl != null
                && Uri.TryCreate(img.BaseUrl.Href, UriKind.Absolute, out var baseUri)
                && !string.Equals(img.BaseUrl.Href, "http://localhost/", StringComparison.OrdinalIgnoreCase)) {
                return baseUri;
            }

            return null;
        }

        private static string ResolveWordImageSource(IHtmlImageElement img, HtmlToWordOptions options) {
            string firstResolved = string.Empty;
            foreach (string candidate in HtmlImageSourceResolver.ResolveImageSourceCandidates(
                         img,
                         ResolveImageBaseUri(img, options),
                         ImageSourceResolutionPolicy)) {
                string resolved = ResolveImageSourcePath(candidate, img, options);
                if (string.IsNullOrWhiteSpace(resolved)) {
                    continue;
                }

                if (string.IsNullOrEmpty(firstResolved)) {
                    firstResolved = resolved;
                }

                if (IsImageSourceAllowedForCurrentMode(resolved, img, options, out _)) {
                    return resolved;
                }
            }

            return firstResolved;
        }

        private static string ResolveImageSourcePath(string source, IHtmlImageElement img, HtmlToWordOptions options) {
            if (string.IsNullOrWhiteSpace(source)
                || source.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)
                || Uri.TryCreate(source, UriKind.Absolute, out _)) {
                return source ?? string.Empty;
            }

            if (!string.IsNullOrEmpty(options.BasePath)) {
                return Path.Combine(options.BasePath, source);
            }

            if (img.BaseUrl != null
                && Uri.TryCreate(img.BaseUrl.Href, UriKind.Absolute, out var baseUri)
                && !string.Equals(img.BaseUrl.Href, "http://localhost/", StringComparison.OrdinalIgnoreCase)) {
                return new Uri(baseUri, source).ToString();
            }

            return source;
        }

        private static HtmlUrlPolicy CreateImageSourceResolutionPolicy() {
            var policy = HtmlUrlPolicy.CreateOfficeIMOProfile();
            policy.DisallowScriptUrls = false;
            return policy;
        }

        private byte[] FetchBytes(Uri uri, HtmlToWordOptions options) {
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
            EnsureImageContentTypeAllowed(response.Content.Headers.ContentType?.MediaType, options);
            var readLimit = GetRemoteImageReadLimit(options);
            var bytes = ReadContentWithLimit(response.Content, readLimit.Limit, readLimit.LimitedByTotalBudget, token);
            ReserveImageBytes(bytes.LongLength, options);
            return bytes;
        }

        private string FetchString(Uri uri, HtmlToWordOptions options, out long reservedBytes) {
            var bytes = FetchBytes(uri, options);
            reservedBytes = bytes.LongLength;
            return Encoding.UTF8.GetString(bytes);
        }

        private (long? Limit, bool LimitedByTotalBudget) GetRemoteImageReadLimit(HtmlToWordOptions options) {
            long? limit = options.MaxImageBytes;
            var limitedByTotalBudget = false;
            if (options.MaxTotalImageBytes.HasValue) {
                var remaining = options.MaxTotalImageBytes.Value - _imageBytesUsed;
                if (remaining <= 0) {
                    throw new HtmlResourceTotalLimitException($"Image resource budget is exhausted; limit is {options.MaxTotalImageBytes.Value} bytes.");
                }

                if (!limit.HasValue || remaining < limit.Value) {
                    limit = remaining;
                    limitedByTotalBudget = true;
                }
            }

            return (limit, limitedByTotalBudget);
        }

        private static byte[] ReadContentWithLimit(HttpContent content, long? maxBytes, bool totalBudgetLimit, CancellationToken cancellationToken) {
            if (maxBytes.HasValue && content.Headers.ContentLength.HasValue && content.Headers.ContentLength.Value > maxBytes.Value) {
                if (totalBudgetLimit) {
                    throw new HtmlResourceTotalLimitException($"Resource length {content.Headers.ContentLength.Value} bytes exceeds remaining total image budget {maxBytes.Value} bytes.");
                }

                throw new HtmlResourceLimitException($"Resource length {content.Headers.ContentLength.Value} bytes exceeds limit {maxBytes.Value} bytes.");
            }

            using var stream = content.ReadAsStreamAsync().GetAwaiter().GetResult();
            using var memory = new MemoryStream();
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                var read = stream.Read(buffer, 0, buffer.Length);
                if (read == 0) {
                    break;
                }

                total += read;
                if (maxBytes.HasValue && total > maxBytes.Value) {
                    if (totalBudgetLimit) {
                        throw new HtmlResourceTotalLimitException($"Resource length exceeded remaining total image budget {maxBytes.Value} bytes.");
                    }

                    throw new HtmlResourceLimitException($"Resource length exceeded limit {maxBytes.Value} bytes.");
                }

                memory.Write(buffer, 0, read);
            }

            return memory.ToArray();
        }

        private static bool IsImageContentTypeAllowed(string? contentType, HtmlToWordOptions options) {
            if (!options.ValidateImageContentTypes || string.IsNullOrWhiteSpace(contentType)) {
                return true;
            }

            var normalized = contentType!.Trim();
            if (options.AllowedImageContentTypes.Contains(normalized)) {
                return true;
            }

            return options.AllowedImageContentTypes.Contains("image/*")
                && normalized.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        }

        private static void EnsureImageContentTypeAllowed(string? contentType, HtmlToWordOptions options) {
            if (!IsImageContentTypeAllowed(contentType, options)) {
                throw new HtmlResourceContentTypeException($"Image content type '{contentType}' is not allowed.");
            }
        }

        private bool TryApplyImageSourcePolicy(string src, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter, WordDocument doc, string alt) {
            if (IsImageSourceAllowed(src, options, out var detail)) {
                return true;
            }

            AddDiagnostic(options, "ImageResourceRejectedByPolicy", "Image resource was skipped because its URI is not allowed by the current image policy.", src, new HtmlResourcePolicyException(detail));
            InsertAltText(currentParagraph, headerFooter, doc, alt);
            return false;
        }

        private static bool IsImageSourceAllowed(string src, HtmlToWordOptions options, out string detail) {
            detail = string.Empty;

            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                return IsImageSchemeAllowed("data", options, out detail);
            }

            if (Uri.TryCreate(src, UriKind.Absolute, out var uri)) {
                if (!IsImageSchemeAllowed(uri.Scheme, options, out detail)) {
                    return false;
                }

                if (!uri.IsFile && options.AllowedImageHosts.Count > 0 && !options.AllowedImageHosts.Contains(uri.Host)) {
                    detail = $"Image host '{uri.Host}' is not allowed.";
                    return false;
                }

                return true;
            }

            if ((File.Exists(src) || IsRootedLocalPath(src)) && !IsImageSchemeAllowed(Uri.UriSchemeFile, options, out detail)) {
                return false;
            }

            return true;
        }

        private static bool IsImageSourceAllowedForCurrentMode(string src, IHtmlImageElement img, HtmlToWordOptions options, out string detail) {
            if (options.ImageProcessing == ImageProcessingMode.EmbedDataUriOnly
                && !src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                detail = "External image was skipped because only data URI images are enabled.";
                return false;
            }

            if (options.ImageProcessing == ImageProcessingMode.LinkExternal
                && !src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)
                && !HasExternalImageDimensionHints(img)) {
                detail = "External image link requires explicit width and height.";
                return false;
            }

            return IsImageSourceAllowed(src, options, out detail);
        }

        private static bool HasExternalImageDimensionHints(IHtmlImageElement img) {
            if (img.DisplayWidth > 0 && img.DisplayHeight > 0) {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(img.GetAttribute("width"))
                && !string.IsNullOrWhiteSpace(img.GetAttribute("height"))) {
                return true;
            }

            string? style = img.GetAttribute("style");
            return !string.IsNullOrWhiteSpace(style)
                   && style!.IndexOf("width", StringComparison.OrdinalIgnoreCase) >= 0
                   && style.IndexOf("height", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsImageSchemeAllowed(string scheme, HtmlToWordOptions options, out string detail) {
            if (options.AllowedImageUriSchemes.Contains(scheme)) {
                detail = string.Empty;
                return true;
            }

            detail = $"Image URI scheme '{scheme}' is not allowed.";
            return false;
        }

        private static bool IsRootedLocalPath(string src) {
            try {
                return Path.IsPathRooted(src);
            } catch (ArgumentException) {
                return false;
            }
        }

        private bool TryReserveImageBytes(long length, HtmlToWordOptions options, string source) {
            try {
                ReserveImageBytes(length, options);
                return true;
            } catch (HtmlResourceTotalLimitException ex) {
                AddDiagnostic(options, "ImageResourceBudgetExceeded", "Image resource exceeded the configured total byte budget and was replaced with alt text when available.", source, ex);
                return false;
            }
        }

        private void ReserveImageBytes(long length, HtmlToWordOptions options) {
            if (!options.MaxTotalImageBytes.HasValue) {
                return;
            }

            var remaining = options.MaxTotalImageBytes.Value - _imageBytesUsed;
            if (length > remaining) {
                throw new HtmlResourceTotalLimitException($"Image resource budget would reach {_imageBytesUsed + length} bytes, exceeding limit {options.MaxTotalImageBytes.Value} bytes.");
            }

            _imageBytesUsed += length;
        }

        private void ReleaseImageBytes(long length, HtmlToWordOptions options) {
            if (length <= 0 || !options.MaxTotalImageBytes.HasValue) {
                return;
            }

            _imageBytesUsed = Math.Max(0, _imageBytesUsed - length);
        }

        private sealed class HtmlResourceLimitException : Exception {
            internal HtmlResourceLimitException(string message) : base(message) {
            }
        }

        private sealed class HtmlResourceTotalLimitException : Exception {
            internal HtmlResourceTotalLimitException(string message) : base(message) {
            }
        }

        private sealed class HtmlResourcePolicyException : Exception {
            internal HtmlResourcePolicyException(string message) : base(message) {
            }
        }

        private sealed class HtmlResourceContentTypeException : Exception {
            internal HtmlResourceContentTypeException(string message) : base(message) {
            }
        }

        private void ProcessSvgElement(AngleSharp.Dom.IElement svg, WordDocument doc, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            double? width = null;
            double? height = null;
            if (double.TryParse(svg.GetAttribute("width")?.Replace("px", string.Empty), out var w)) width = w;
            if (double.TryParse(svg.GetAttribute("height")?.Replace("px", string.Empty), out var h)) height = h;

            var paragraph = currentParagraph ?? (headerFooter != null ? headerFooter.AddParagraph() : section.AddParagraph());
            long reservedBytes = 0;
            try {
                var svgByteCount = Encoding.UTF8.GetByteCount(svg.OuterHtml);
                if (options.MaxImageBytes.HasValue && svgByteCount > options.MaxImageBytes.Value) {
                    AddDiagnostic(options, "ImageResourceTooLarge", "Inline SVG exceeded the configured byte limit and was skipped.", "svg");
                    return;
                }
                if (!TryReserveImageBytes(svgByteCount, options, "svg")) {
                    return;
                }
                reservedBytes = svgByteCount;

                SvgHelper.AddSvg(paragraph, svg.OuterHtml, width, height, string.Empty);
                reservedBytes = 0;
            } catch (Exception ex) {
                ReleaseImageBytes(reservedBytes, options);
                AddDiagnostic(options, "InlineSvgEmbedFailed", "Inline SVG could not be embedded and was skipped.", "svg", ex);
            }
        }
    }
}
