using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using System.Globalization;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static readonly string[] WordImageSrcSetAttributes = { "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset" };
        private static readonly string[] WordPictureSourceAttributes = { "src", "data-src", "data-original", "data-original-src", "data-lazy-src" };
        private static readonly string[] WordImageLazySourceAttributes = { "data-src", "data-original", "data-original-src", "data-lazy-src" };
        private static readonly string[] WordImageSourceAttributes = { "src" };

        private void ProcessImage(IHtmlImageElement img, WordDocument doc, HtmlToWordOptions options, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            var src = ResolveWordImageSource(img, options);
            if (string.IsNullOrEmpty(src)) {
                if (HasImageSourceCandidateAttribute(img)) {
                    string altText = img.AlternativeText ?? string.Empty;
                    AddDiagnostic(options, "ImageResourceRejectedByPolicy", "Image resource candidates were skipped because their URIs are not allowed by the current image policy.", "responsive image candidates");
                    InsertAltText(currentParagraph, headerFooter, doc, altText);
                }

                return;
            }
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
                    if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
                        AddDiagnostic(options, "ImageDataUriInvalid", "Image data URI could not be decoded or embedded and was replaced with alt text when available.", src);
                        return false;
                    }
                    if (options.MaxImageBytes.HasValue && bytes.LongLength > options.MaxImageBytes.Value) {
                        AddDiagnostic(options, "ImageResourceTooLarge", "Image data URI exceeded the configured byte limit and was replaced with alt text when available.", "data:image");
                        return false;
                    }
                    if (!TryReserveImageBytes(bytes.LongLength, options, "data:image")) {
                        return false;
                    }
                    reservedBytes = bytes.LongLength;
                    paragraph ??= headerFooter != null ? headerFooter.AddParagraph() : doc.AddParagraph();
                    using var imageStream = new MemoryStream(bytes);
                    paragraph.AddImage(imageStream, "image." + ext, width, height, wrap, description: alt);
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

            if (TryGetUsableDocumentBaseUri(img.BaseUrl?.Href, out var baseUri)) {
                return baseUri;
            }

            return null;
        }

        private string ResolveWordImageSource(IHtmlImageElement img, HtmlToWordOptions options) {
            string firstResolved = string.Empty;
            int remoteCandidateProbeCount = 0;
            foreach (string candidate in EnumerateWordImageSourceCandidates(img, options)) {
                string resolved = ResolveImageSourcePath(candidate, img, options);
                if (string.IsNullOrWhiteSpace(resolved)) {
                    continue;
                }

                if (IsUnresolvedRelativeImageSource(resolved)) {
                    if (string.IsNullOrEmpty(firstResolved)) {
                        firstResolved = resolved;
                    }

                    continue;
                }

                if (IsImageSourceAllowedForCurrentMode(resolved, img, options, out _)) {
                    if (IsRemoteEmbeddedImageSource(resolved, options)) {
                        if (!CanProbeRemoteImageCandidate(options, remoteCandidateProbeCount)) {
                            if (string.IsNullOrEmpty(firstResolved)) {
                                firstResolved = resolved;
                            }

                            continue;
                        }

                        remoteCandidateProbeCount++;
                        if (!TryFetchRemoteImageCandidate(resolved, options)) {
                            if (string.IsNullOrEmpty(firstResolved)) {
                                firstResolved = resolved;
                            }

                            continue;
                        }
                    }

                    if (IsLocalEmbeddedImageSource(resolved, options) && !TryProbeLocalImageCandidate(resolved, options)) {
                        if (string.IsNullOrEmpty(firstResolved)) {
                            firstResolved = resolved;
                        }

                        continue;
                    }

                    return resolved;
                }

                if (string.IsNullOrEmpty(firstResolved)) {
                    firstResolved = resolved;
                }
            }

            return firstResolved;
        }

        private static bool CanProbeRemoteImageCandidate(HtmlToWordOptions options, int probeCount) {
            return !options.MaxRemoteImageCandidateProbes.HasValue
                || probeCount < options.MaxRemoteImageCandidateProbes.Value;
        }

        private static IEnumerable<string> EnumerateWordImageSourceCandidates(IHtmlImageElement img, HtmlToWordOptions options) {
            Uri? baseUri = ResolveImageBaseUri(img, options);
            var responsiveCandidateState = new ResponsiveImageCandidateState(options);
            if (img.ParentElement != null
                && img.ParentElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
                foreach (var child in img.ParentElement.Children) {
                    if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    if (!IsImageContentTypeAllowed(child.GetAttribute("type"), options)) {
                        continue;
                    }

                    foreach (string candidate in ResolveImageCandidatesFromSourceElement(child, img, baseUri, options, responsiveCandidateState)) {
                        yield return candidate;
                        if (responsiveCandidateState.HasReachedAnyLimit) {
                            break;
                        }
                    }

                    if (responsiveCandidateState.HasReachedAnyLimit) {
                        break;
                    }
                }
            }

            foreach (string candidate in ResolveImageUrlAttributeCandidates(img, baseUri, WordImageLazySourceAttributes, options, responsiveCandidateState)) {
                yield return candidate;
            }

            foreach (string candidate in ResolveImageSrcSetCandidates(img, baseUri, options, responsiveCandidateState)) {
                yield return candidate;
                if (responsiveCandidateState.HasReachedAnyLimit) {
                    break;
                }
            }

            foreach (string candidate in ResolveImageUrlAttributeCandidates(img, baseUri, WordImageSourceAttributes, options, responsiveCandidateState)) {
                yield return candidate;
            }
        }

        private static IEnumerable<string> ResolveImageCandidatesFromSourceElement(AngleSharp.Dom.IElement sourceElement, IHtmlImageElement img, Uri? baseUri, HtmlToWordOptions options, ResponsiveImageCandidateState state) {
            foreach (string attributeName in WordImageSrcSetAttributes) {
                foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(sourceElement.GetAttribute(attributeName))) {
                    string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, options.ResourceUrlPolicy);
                    if (!IsResolvedImageCandidateAllowedForEnumeration(resolved, img, options)) {
                        state.TrackRejectedResponsiveCandidate();
                        if (state.HasReachedScanLimit) {
                            yield break;
                        }

                        continue;
                    }

                    if (!state.TryTrackResponsiveCandidate(resolved, out string tracked)) {
                        if (state.HasReachedAnyLimit) {
                            yield break;
                        }

                        continue;
                    }

                    yield return tracked;
                    if (state.HasReachedAnyLimit) {
                        yield break;
                    }
                }
            }

            foreach (string attributeName in WordPictureSourceAttributes) {
                if (state.HasReachedAnyLimit) {
                    yield break;
                }

                string? rawValue = sourceElement.GetAttribute(attributeName);
                if (string.IsNullOrWhiteSpace(rawValue)) {
                    continue;
                }

                string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(rawValue, baseUri, options.ResourceUrlPolicy);
                if (!IsResolvedImageCandidateAllowedForEnumeration(resolved, img, options)) {
                    state.TrackRejectedResponsiveCandidate();
                    continue;
                }

                if (state.TryTrackResponsiveCandidate(resolved, out string tracked)) {
                    yield return tracked;
                }
            }
        }

        private static IEnumerable<string> ResolveImageSrcSetCandidates(IHtmlImageElement img, Uri? baseUri, HtmlToWordOptions options, ResponsiveImageCandidateState state) {
            foreach (string attributeName in WordImageSrcSetAttributes) {
                foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(img.GetAttribute(attributeName))) {
                    string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, options.ResourceUrlPolicy);
                    if (!IsResolvedImageCandidateAllowedForEnumeration(resolved, img, options)) {
                        state.TrackRejectedResponsiveCandidate();
                        if (state.HasReachedScanLimit) {
                            yield break;
                        }

                        continue;
                    }

                    if (!state.TryTrackResponsiveCandidate(resolved, out string tracked)) {
                        if (state.HasReachedAnyLimit) {
                            yield break;
                        }

                        continue;
                    }

                    yield return tracked;
                    if (state.HasReachedAnyLimit) {
                        yield break;
                    }
                }
            }
        }

        private static IEnumerable<string> ResolveImageUrlAttributeCandidates(IHtmlImageElement img, Uri? baseUri, IEnumerable<string> attributeNames, HtmlToWordOptions options, ResponsiveImageCandidateState state) {
            foreach (string attributeName in attributeNames) {
                string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(img.GetAttribute(attributeName), baseUri, options.ResourceUrlPolicy);
                if (state.TryTrackFixedCandidate(resolved, out string tracked)) {
                    yield return tracked;
                }
            }
        }

        private static bool IsResolvedImageCandidateAllowedForEnumeration(string? resolved, IHtmlImageElement img, HtmlToWordOptions options) {
            if (string.IsNullOrWhiteSpace(resolved)) {
                return false;
            }

            string source = resolved!;
            if (options.ImageProcessing == ImageProcessingMode.EmbedDataUriOnly
                && !source.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (options.ImageProcessing == ImageProcessingMode.LinkExternal
                && !source.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)
                && !HasExternalImageDimensionHints(img)) {
                return false;
            }

            if (!IsImageSourceAllowed(source, options, out _)) {
                return false;
            }

            return !source.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
                || source.StartsWith("data:image", StringComparison.OrdinalIgnoreCase);
        }

        private sealed class ResponsiveImageCandidateState {
            private readonly HtmlToWordOptions _options;
            private readonly HashSet<string> _seen = new HashSet<string>(StringComparer.Ordinal);
            private int _count;
            private int _scanned;

            internal ResponsiveImageCandidateState(HtmlToWordOptions options) {
                _options = options;
            }

            internal bool HasReachedLimit => _options.MaxImageSourceCandidates.HasValue
                && _count >= _options.MaxImageSourceCandidates.Value;

            internal bool HasReachedScanLimit => _options.MaxImageSourceCandidates.HasValue
                && _scanned >= GetResponsiveCandidateScanLimit(_options.MaxImageSourceCandidates.Value);

            internal bool HasReachedAnyLimit => HasReachedLimit || HasReachedScanLimit;

            internal void TrackRejectedResponsiveCandidate() {
                _scanned++;
            }

            internal bool TryTrackResponsiveCandidate(string? candidate, out string tracked) {
                tracked = string.Empty;
                _scanned++;
                if (string.IsNullOrWhiteSpace(candidate) || HasReachedLimit || !_seen.Add(candidate!)) {
                    return false;
                }

                _count++;
                tracked = candidate!;
                return true;
            }

            internal bool TryTrackFixedCandidate(string? candidate, out string tracked) {
                tracked = string.Empty;
                if (string.IsNullOrWhiteSpace(candidate) || !_seen.Add(candidate!)) {
                    return false;
                }

                tracked = candidate!;
                return true;
            }

            private static long GetResponsiveCandidateScanLimit(int maxCandidates) {
                return Math.Max((long)maxCandidates, (long)maxCandidates * 4L);
            }
        }

        private static bool HasImageSourceCandidateAttribute(IHtmlImageElement img) {
            if (HasAnyAttribute(img, WordImageLazySourceAttributes)
                || HasAnyAttribute(img, WordImageSrcSetAttributes)
                || HasAnyAttribute(img, WordImageSourceAttributes)) {
                return true;
            }

            if (img.ParentElement == null
                || !img.ParentElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            foreach (var child in img.ParentElement.Children) {
                if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (HasAnyAttribute(child, WordImageSrcSetAttributes)
                    || HasAnyAttribute(child, WordPictureSourceAttributes)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasAnyAttribute(AngleSharp.Dom.IElement element, IEnumerable<string> attributeNames) {
            foreach (string attributeName in attributeNames) {
                if (!string.IsNullOrWhiteSpace(element.GetAttribute(attributeName))) {
                    return true;
                }
            }

            return false;
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

            if (TryGetUsableDocumentBaseUri(img.BaseUrl?.Href, out var baseUri)) {
                return new Uri(baseUri, source).ToString();
            }

            return source;
        }

        private static bool IsUnresolvedRelativeImageSource(string source) {
            if (string.IsNullOrWhiteSpace(source)
                || source.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)
                || Uri.TryCreate(source, UriKind.Absolute, out _)
                || File.Exists(source)) {
                return false;
            }

            try {
                return !Path.IsPathRooted(source);
            } catch (ArgumentException) {
                return true;
            }
        }

        private static bool TryGetUsableDocumentBaseUri(string? value, out Uri baseUri) {
            baseUri = null!;
            if (string.IsNullOrWhiteSpace(value)
                || string.Equals(value, "http://localhost/", StringComparison.OrdinalIgnoreCase)
                || !Uri.TryCreate(value, UriKind.Absolute, out Uri? candidate)
                || string.Equals(candidate.Scheme, "about", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            baseUri = candidate;
            return true;
        }

        private async Task PrefetchRemoteImagesAsync(
            IHtmlDocument document,
            HtmlToWordOptions options,
            CancellationToken cancellationToken) {
            foreach (IElement element in document.QuerySelectorAll("img")) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!(element is IHtmlImageElement image)) {
                    continue;
                }

                int remoteCandidateProbeCount = 0;
                foreach (string candidate in EnumerateWordImageSourceCandidates(image, options)) {
                    string resolved = ResolveImageSourcePath(candidate, image, options);
                    if (!IsRemoteEmbeddedImageSource(resolved, options)
                        || !IsImageSourceAllowedForCurrentMode(resolved, image, options, out _)) {
                        continue;
                    }

                    if (!CanProbeRemoteImageCandidate(options, remoteCandidateProbeCount)) {
                        continue;
                    }

                    remoteCandidateProbeCount++;
                    await PrefetchRemoteImageCandidateAsync(resolved, options, cancellationToken).ConfigureAwait(false);
                }
            }
        }

        private async Task PrefetchRemoteImageCandidateAsync(
            string source,
            HtmlToWordOptions options,
            CancellationToken cancellationToken) {
            if (!Uri.TryCreate(source, UriKind.Absolute, out Uri? uri)
                || _remoteImageBytesCache.ContainsKey(uri.AbsoluteUri)
                || _remoteImageFailureCache.ContainsKey(uri.AbsoluteUri)) {
                return;
            }

            try {
                using var cts = _resourceTimeout.HasValue
                    ? CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                    : null;
                CancellationToken token = cts?.Token ?? cancellationToken;
                if (cts != null && _resourceTimeout.HasValue) {
                    cts.CancelAfter(_resourceTimeout.Value);
                }

                using var request = new HttpRequestMessage(HttpMethod.Get, uri);
                using var response = await _httpClient.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead,
                    token).ConfigureAwait(false);
                response.EnsureSuccessStatusCode();
                EnsureImageContentTypeAllowed(response.Content.Headers.ContentType?.MediaType, options);
                (long? limit, bool totalBudgetLimit) = GetRemoteImagePrefetchReadLimit(options);
                byte[] bytes = await ReadContentWithLimitAsync(
                    response.Content,
                    limit,
                    totalBudgetLimit,
                    token).ConfigureAwait(false);
                if (!IsEmbeddableImageData(bytes, out string detail)) {
                    throw new InvalidDataException(detail);
                }

                _remoteImageBytesCache[uri.AbsoluteUri] = bytes;
                _remoteImageBytesFetched += bytes.LongLength;
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                throw;
            } catch (Exception ex) {
                _remoteImageFailureCache[uri.AbsoluteUri] = ex;
            }
        }

        private byte[] FetchBytes(Uri uri, HtmlToWordOptions options) {
            string cacheKey = uri.AbsoluteUri;
            if (_remoteImageBytesCache.TryGetValue(cacheKey, out byte[]? cachedBytes)) {
                ReserveImageBytes(cachedBytes.LongLength, options);
                return cachedBytes;
            }

            if (_remoteImageFailureCache.TryGetValue(cacheKey, out Exception? cachedFailure)) {
                throw cachedFailure;
            }

            throw new InvalidOperationException("Remote image resources must be prepared by the asynchronous conversion pipeline before document projection.");
        }

        private bool TryFetchRemoteImageCandidate(string source, HtmlToWordOptions options) {
            if (!Uri.TryCreate(source, UriKind.Absolute, out var uri)) {
                return false;
            }

            long reservedBytes = 0;
            try {
                byte[] bytes = FetchBytes(uri, options);
                reservedBytes = bytes.LongLength;
                if (!IsEmbeddableImageData(bytes, out var detail)) {
                    _remoteImageBytesCache.Remove(uri.AbsoluteUri);
                    _remoteImageFailureCache[uri.AbsoluteUri] = new InvalidDataException(detail);
                    return false;
                }

                _remoteImageBytesCache[uri.AbsoluteUri] = bytes;
                _remoteImageFailureCache.Remove(uri.AbsoluteUri);
                return true;
            } catch (OperationCanceledException ex) when (!_cancellationToken.IsCancellationRequested) {
                _remoteImageFailureCache[uri.AbsoluteUri] = ex;
                return false;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                _remoteImageFailureCache[uri.AbsoluteUri] = ex;
                return false;
            } finally {
                ReleaseImageBytes(reservedBytes, options);
            }
        }

        private string FetchString(Uri uri, HtmlToWordOptions options, out long reservedBytes) {
            var bytes = FetchBytes(uri, options);
            reservedBytes = bytes.LongLength;
            return Encoding.UTF8.GetString(bytes);
        }

        private (long? Limit, bool LimitedByTotalBudget) GetRemoteImagePrefetchReadLimit(HtmlToWordOptions options) {
            long? limit = options.MaxImageBytes;
            bool limitedByTotalBudget = false;
            if (options.MaxTotalImageBytes.HasValue) {
                long remaining = options.MaxTotalImageBytes.Value - _remoteImageBytesFetched;
                if (remaining <= 0) {
                    throw new HtmlResourceTotalLimitException($"Remote image fetch budget is exhausted; limit is {options.MaxTotalImageBytes.Value} bytes.");
                }

                if (!limit.HasValue || remaining < limit.Value) {
                    limit = remaining;
                    limitedByTotalBudget = true;
                }
            }

            return (limit, limitedByTotalBudget);
        }

        private static async Task<byte[]> ReadContentWithLimitAsync(
            HttpContent content,
            long? maxBytes,
            bool totalBudgetLimit,
            CancellationToken cancellationToken) {
            if (maxBytes.HasValue && content.Headers.ContentLength.HasValue && content.Headers.ContentLength.Value > maxBytes.Value) {
                if (totalBudgetLimit) {
                    throw new HtmlResourceTotalLimitException($"Resource length {content.Headers.ContentLength.Value} bytes exceeds remaining total image budget {maxBytes.Value} bytes.");
                }

                throw new HtmlResourceLimitException($"Resource length {content.Headers.ContentLength.Value} bytes exceeds limit {maxBytes.Value} bytes.");
            }

            using var stream = await content.ReadAsStreamAsync().ConfigureAwait(false);
            using var memory = new MemoryStream();
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
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

            var normalized = NormalizeImageContentType(contentType!);
            if (options.AllowedImageContentTypes.Contains(normalized)) {
                return true;
            }

            return options.AllowedImageContentTypes.Contains("image/*")
                && normalized.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeImageContentType(string contentType) {
            var normalized = contentType.Trim();
            int parameterIndex = normalized.IndexOf(';');
            if (parameterIndex >= 0) {
                normalized = normalized.Substring(0, parameterIndex).Trim();
            }

            return normalized;
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

        private bool IsImageSourceAllowedForCurrentMode(string src, IHtmlImageElement img, HtmlToWordOptions options, out string detail) {
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

            if (!IsImageSourceAllowed(src, options, out detail)) {
                return false;
            }

            if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
                && !src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                detail = "Only image data URI candidates are supported.";
                return false;
            }

            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                return IsDataImageCandidateAllowed(src, options, out detail);
            }

            return true;
        }

        private bool IsDataImageCandidateAllowed(string src, HtmlToWordOptions options, out string detail) {
            detail = string.Empty;
            if (!HtmlImageDataUri.TryParse(src, out var dataUri)) {
                detail = "Image data URI could not be parsed.";
                return false;
            }

            if (!IsImageContentTypeAllowed(dataUri.MediaType, options)) {
                detail = $"Image data URI content type '{dataUri.MediaType}' is not allowed.";
                return false;
            }

            try {
                long estimatedBytes;
                if (dataUri.IsBase64) {
                    if (_imageCache.ContainsKey(src)) {
                        return true;
                    }

                    if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
                        detail = "Image data URI could not be decoded.";
                        return false;
                    }

                    if (!TryIdentifyEmbeddableImageData(bytes, out var imageInfo, out detail)) {
                        return false;
                    }

                    if (dataUri.MediaType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase)
                        && !IsEmbeddableSvgBytes(bytes, out detail)) {
                        return false;
                    }

                    if (!IsIdentifiedImageContentTypeAllowed(imageInfo, options, out detail)) {
                        return false;
                    }

                    estimatedBytes = bytes.LongLength;
                } else {
                    if (!dataUri.MediaType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                        detail = "Only SVG text data URI images are supported when the payload is not base64 encoded.";
                        return false;
                    }

                    string svgText = dataUri.DecodeText();
                    if (!IsEmbeddableSvgText(svgText, out detail)) {
                        return false;
                    }

                    estimatedBytes = Encoding.UTF8.GetByteCount(svgText);
                }

                if (options.MaxImageBytes.HasValue && estimatedBytes > options.MaxImageBytes.Value) {
                    detail = $"Image data URI estimated size {estimatedBytes} bytes exceeds limit {options.MaxImageBytes.Value} bytes.";
                    return false;
                }

                if (options.MaxTotalImageBytes.HasValue && estimatedBytes > options.MaxTotalImageBytes.Value - _imageBytesUsed) {
                    detail = $"Image data URI estimated size {estimatedBytes} bytes exceeds remaining image byte budget.";
                    return false;
                }

                return true;
            } catch (UriFormatException ex) {
                detail = ex.Message;
                return false;
            } catch (FormatException ex) {
                detail = ex.Message;
                return false;
            }
        }

        private static bool IsEmbeddableImageData(byte[] bytes, out string detail) {
            return TryIdentifyEmbeddableImageData(bytes, out _, out detail);
        }

        private static bool TryIdentifyEmbeddableImageData(byte[] bytes, out OfficeImageInfo imageInfo, out string detail) {
            detail = string.Empty;
            if (!OfficeImageReader.TryIdentifyByContent(bytes, null, out imageInfo)) {
                detail = "Image data URI payload is not a supported image.";
                return false;
            }

            return true;
        }

        private static bool IsIdentifiedImageContentTypeAllowed(OfficeImageInfo imageInfo, HtmlToWordOptions options, out string detail) {
            detail = string.Empty;
            if (IsImageContentTypeAllowed(imageInfo.MimeType, options)) {
                return true;
            }

            if (imageInfo.Format == OfficeImageFormat.Jpeg && IsImageContentTypeAllowed("image/jpg", options)) {
                return true;
            }

            detail = $"Image payload content type '{imageInfo.MimeType}' is not allowed.";
            return false;
        }

        private static bool IsEmbeddableSvgText(string svgText, out string detail) {
            detail = string.Empty;
            byte[] bytes = Encoding.UTF8.GetBytes(svgText);
            return IsEmbeddableSvgBytes(bytes, out detail);
        }

        private static bool IsEmbeddableSvgBytes(byte[] bytes, out string detail) {
            detail = string.Empty;
            if (!OfficeImageReader.TryIdentifyByContent(bytes, null, out var info) || info.Format != OfficeImageFormat.Svg) {
                detail = "SVG data URI payload is not a valid SVG image.";
                return false;
            }

            return true;
        }

        private bool TryProbeLocalImageCandidate(string source, HtmlToWordOptions options) {
            if (!TryGetLocalImagePath(source, out string path)) {
                return false;
            }

            long reservedBytes = 0;
            try {
                reservedBytes = EnsureFileWithinImageLimits(path, options);
                byte[] bytes = File.ReadAllBytes(path);
                if (!IsEmbeddableImageData(bytes, out _)) {
                    return false;
                }

                return true;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception) {
                return false;
            } finally {
                ReleaseImageBytes(reservedBytes, options);
            }
        }

        private static bool IsRemoteEmbeddedImageSource(string source, HtmlToWordOptions options) {
            if (options.ImageProcessing == ImageProcessingMode.LinkExternal) {
                return false;
            }

            return Uri.TryCreate(source, UriKind.Absolute, out var uri)
                   && !uri.IsFile
                   && (uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                       || uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsLocalEmbeddedImageSource(string source, HtmlToWordOptions options) {
            if (options.ImageProcessing == ImageProcessingMode.LinkExternal) {
                return false;
            }

            return TryGetLocalImagePath(source, out _);
        }

        private static bool TryGetLocalImagePath(string source, out string path) {
            path = string.Empty;
            if (Uri.TryCreate(source, UriKind.Absolute, out var uri)) {
                if (!uri.IsFile) {
                    return false;
                }

                path = uri.LocalPath;
                return true;
            }

            if (!File.Exists(source)) {
                return false;
            }

            path = source;
            return true;
        }

        private static bool HasExternalImageDimensionHints(IHtmlImageElement img) {
            if (img.DisplayWidth > 0 && img.DisplayHeight > 0) {
                return true;
            }

            return TryParsePixelValue(img.GetAttribute("width")).HasValue
                   && TryParsePixelValue(img.GetAttribute("height")).HasValue;
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
