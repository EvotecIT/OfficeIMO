using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static bool LocalPathAllowed(string path, MarkdownToWordOptions options) {
            if (!options.AllowLocalImages) return false;
            if (options.AllowedImageDirectories.Count == 0) return true;
            try {
                var full = System.IO.Path.GetFullPath(path);
                foreach (var root in options.AllowedImageDirectories) {
                    var rootFull = System.IO.Path.GetFullPath(root.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) + System.IO.Path.DirectorySeparatorChar);
                    if (full.StartsWith(rootFull, System.StringComparison.OrdinalIgnoreCase)) return true;
                }
            } catch { return false; }
            return false;
        }

        private static double EstimatePageContentWidthPixels(WordDocument document) {
            var section = document.Sections.FirstOrDefault();
            var pageWidthTwips = (double?)section?.PageSettings?.Width?.Value ?? DefaultPageWidthTwips;
            var leftMarginTwips = (double?)section?.Margins?.Left?.Value ?? DefaultHorizontalMarginTwips;
            var rightMarginTwips = (double?)section?.Margins?.Right?.Value ?? DefaultHorizontalMarginTwips;
            var contentTwips = pageWidthTwips - leftMarginTwips - rightMarginTwips;

            if (contentTwips < MinimumContentWidthPixels) {
                contentTwips = DefaultPageWidthTwips - (DefaultHorizontalMarginTwips * 2);
            }

            if (contentTwips < MinimumContentWidthPixels) {
                return MinimumContentWidthPixels;
            }

            return contentTwips * PixelsPerInch / TwipsPerInch;
        }

        private static System.Net.Http.HttpClient CreateRemoteImageClient(TimeSpan timeout, bool bypassProxy = false) {
            System.Net.Http.HttpClient client;
            if (bypassProxy) {
                var handler = new System.Net.Http.HttpClientHandler {
                    Proxy = null,
                    UseProxy = false
                };
                client = new System.Net.Http.HttpClient(handler, disposeHandler: true);
            } else {
                client = new System.Net.Http.HttpClient();
            }

            client.Timeout = timeout;
            client.DefaultRequestHeaders.UserAgent.ParseAdd("OfficeIMO.Word.Markdown");
            return client;
        }

        private static bool IsLoopbackImageUri(Uri uri) {
            if (uri == null) {
                return false;
            }

            if (uri.IsLoopback) {
                return true;
            }

            return string.Equals(uri.Host, "localhost", StringComparison.OrdinalIgnoreCase);
        }


        private static TimeSpan ResolveRemoteImageTimeout(MarkdownToWordOptions options) {
            if (options.RemoteImageDownloadTimeout <= TimeSpan.Zero) {
                return DefaultRemoteImageDownloadTimeout;
            }

            return options.RemoteImageDownloadTimeout;
        }


        private static byte[] DownloadRemoteImageBytes(Uri uri, MarkdownToWordOptions options) {
            var timeout = ResolveRemoteImageTimeout(options);
            // Fresh clients avoid stale loopback/proxy behavior on older framework handlers.
            using var client = CreateRemoteImageClient(timeout, bypassProxy: IsLoopbackImageUri(uri));
            return client.GetByteArrayAsync(uri).GetAwaiter().GetResult();
        }

        private static void RenderMarkdownImageIntoParagraph(
            WordParagraph paragraph,
            string pathOrUrl,
            string? alt,
            double? requestedWidth,
            double? requestedHeight,
            MarkdownToWordOptions options,
            double pageContentWidthPixels,
            int listLevel,
            int quoteDepth,
            string contextPrefix) {
            var imageSource = pathOrUrl ?? string.Empty;
            var contextWidthLimit = ResolveContextWidthLimitPixels(options.ImageLayout, pageContentWidthPixels, listLevel, quoteDepth);

            if (TryRenderDataUriImage(paragraph, imageSource, alt, requestedWidth, requestedHeight, options, pageContentWidthPixels, contextWidthLimit, contextPrefix)) {
                return;
            }

            if (System.IO.File.Exists(imageSource)) {
                if (options.AllowLocalImages && LocalPathAllowed(imageSource, options)) {
                    double? naturalW = null;
                    double? naturalH = null;
                    if (TryGetImageDimensionsFromFile(imageSource, out var fileW, out var fileH)) {
                        naturalW = fileW;
                        naturalH = fileH;
                    }

                    ResolveImageDimensions(
                        options,
                        source: imageSource,
                        context: contextPrefix + "-local",
                        requestedWidth: requestedWidth,
                        requestedHeight: requestedHeight,
                        naturalWidth: naturalW,
                        naturalHeight: naturalH,
                        pageContentWidthPixels: pageContentWidthPixels,
                        contextWidthLimitPixels: contextWidthLimit,
                        out var finalW,
                        out var finalH,
                        out _);

                    paragraph.AddImage(imageSource, finalW, finalH, description: alt ?? string.Empty);
                } else {
                    AddImageFallbackText(paragraph, alt, System.IO.Path.GetFileName(imageSource), options);
                }

                return;
            }

            if (System.Uri.TryCreate(imageSource, System.UriKind.Absolute, out var uri)) {
                if (options.AllowedImageSchemes.Contains(uri.Scheme) &&
                    (options.ImageUrlValidator == null || options.ImageUrlValidator(uri))) {
                    if (options.AllowRemoteImages) {
                        try {
                            var bytes = DownloadRemoteImageBytes(uri, options);
                            var fileName = System.IO.Path.GetFileName(uri.LocalPath);
                            if (string.IsNullOrWhiteSpace(fileName)) {
                                fileName = "image";
                            }

                            double? naturalW = null;
                            double? naturalH = null;
                            if (TryGetImageDimensionsFromBytes(bytes, out var remoteW, out var remoteH)) {
                                naturalW = remoteW;
                                naturalH = remoteH;
                            }

                            ResolveImageDimensions(
                                options,
                                source: uri.ToString(),
                                context: contextPrefix + "-remote",
                                requestedWidth: requestedWidth,
                                requestedHeight: requestedHeight,
                                naturalWidth: naturalW,
                                naturalHeight: naturalH,
                                pageContentWidthPixels: pageContentWidthPixels,
                                contextWidthLimitPixels: contextWidthLimit,
                                out var finalW,
                                out var finalH,
                                out _);

                            using var stream = new System.IO.MemoryStream(bytes, writable: false);
                            paragraph.AddImage(stream, fileName, finalW, finalH, description: alt ?? string.Empty);
                        } catch (Exception ex) {
                            options.OnWarning?.Invoke($"Remote image '{uri}' could not be downloaded. {ex.Message}");
                            if (options.FallbackRemoteImagesToHyperlinks) {
                                paragraph.AddHyperLink(alt ?? uri.ToString(), uri);
                            }
                        }
                    } else if (options.FallbackRemoteImagesToHyperlinks) {
                        paragraph.AddHyperLink(alt ?? uri.ToString(), uri);
                    }
                } else if (options.FallbackRemoteImagesToHyperlinks) {
                    paragraph.AddHyperLink(alt ?? uri.ToString(), uri);
                }

                return;
            }

            AddImageFallbackText(paragraph, alt, imageSource, options);
        }

        private static bool TryRenderDataUriImage(
            WordParagraph paragraph,
            string source,
            string? alt,
            double? requestedWidth,
            double? requestedHeight,
            MarkdownToWordOptions options,
            double pageContentWidthPixels,
            double? contextWidthLimit,
            string contextPrefix) {
            if (!source.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (!options.AllowDataUriImages) {
                options.OnWarning?.Invoke("Data URI image was skipped because AllowDataUriImages is disabled.");
                AddImageFallbackText(paragraph, alt, string.Empty, options);
                return true;
            }

            if (!TryDecodeDataUriImage(source, options, out var bytes, out var fileName, out var warning)) {
                if (!string.IsNullOrWhiteSpace(warning)) {
                    options.OnWarning?.Invoke(warning!);
                }

                AddImageFallbackText(paragraph, alt, string.Empty, options);
                return true;
            }

            double? naturalW = null;
            double? naturalH = null;
            if (TryGetImageDimensionsFromBytes(bytes, out var dataW, out var dataH)) {
                naturalW = dataW;
                naturalH = dataH;
            }

            ResolveImageDimensions(
                options,
                source: "data:image",
                context: contextPrefix + "-data-uri",
                requestedWidth: requestedWidth,
                requestedHeight: requestedHeight,
                naturalWidth: naturalW,
                naturalHeight: naturalH,
                pageContentWidthPixels: pageContentWidthPixels,
                contextWidthLimitPixels: contextWidthLimit,
                out var finalW,
                out var finalH,
                out _);

            using var stream = new System.IO.MemoryStream(bytes, writable: false);
            paragraph.AddImage(stream, fileName, finalW, finalH, description: alt ?? string.Empty);
            return true;
        }

        private static bool TryDecodeDataUriImage(
            string source,
            MarkdownToWordOptions options,
            out byte[] bytes,
            out string fileName,
            out string? warning) {
            bytes = Array.Empty<byte>();
            fileName = "image.png";
            warning = null;

            int commaIndex = source.IndexOf(',');
            if (commaIndex < 0) {
                warning = "Data URI image is missing a payload separator.";
                return false;
            }

            string metadata = source.Substring(5, commaIndex - 5);
            string payload = source.Substring(commaIndex + 1);
            var metadataParts = metadata.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string contentType = metadataParts.Length > 0 ? metadataParts[0].Trim() : string.Empty;
            if (!contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) {
                warning = $"Data URI content type '{contentType}' is not an image.";
                return false;
            }

            bool base64 = metadataParts.Any(part => string.Equals(part.Trim(), "base64", StringComparison.OrdinalIgnoreCase));
            if (!base64) {
                warning = "Data URI image payload must be base64 encoded.";
                return false;
            }

            if (options.MaxDataUriImageBytes >= 0 &&
                TryEstimateBase64DecodedLength(payload, out long estimatedBytes) &&
                estimatedBytes > options.MaxDataUriImageBytes) {
                warning = $"Data URI image payload is at least {estimatedBytes} bytes, exceeding the configured limit of {options.MaxDataUriImageBytes} bytes.";
                return false;
            }

            try {
                bytes = System.Convert.FromBase64String(payload);
            } catch (FormatException ex) {
                warning = $"Data URI image payload could not be decoded. {ex.Message}";
                return false;
            }

            if (options.MaxDataUriImageBytes >= 0 && bytes.LongLength > options.MaxDataUriImageBytes) {
                warning = $"Data URI image payload is {bytes.LongLength} bytes, exceeding the configured limit of {options.MaxDataUriImageBytes} bytes.";
                bytes = Array.Empty<byte>();
                return false;
            }

            fileName = "image" + ResolveImageExtension(contentType);
            return true;
        }

        private static bool TryEstimateBase64DecodedLength(string payload, out long decodedLength) {
            decodedLength = 0;
            long meaningfulLength = 0;
            int padding = 0;
            bool countingPadding = true;
            for (int i = payload.Length - 1; i >= 0 && countingPadding; i--) {
                char c = payload[i];
                if (char.IsWhiteSpace(c)) {
                    continue;
                }

                if (c == '=' && padding < 2) {
                    padding++;
                    continue;
                }

                countingPadding = false;
            }

            for (int i = 0; i < payload.Length; i++) {
                if (!char.IsWhiteSpace(payload[i])) {
                    meaningfulLength++;
                }
            }

            if (meaningfulLength == 0) {
                return true;
            }

            long groups = (meaningfulLength + 3) / 4;
            if (groups > long.MaxValue / 3) {
                decodedLength = long.MaxValue;
                return true;
            }

            decodedLength = groups * 3 - padding;
            if (decodedLength < 0) {
                decodedLength = 0;
            }

            return true;
        }

        private static string ResolveImageExtension(string contentType) {
            return contentType.ToLowerInvariant() switch {
                "image/jpeg" => ".jpg",
                "image/jpg" => ".jpg",
                "image/gif" => ".gif",
                "image/bmp" => ".bmp",
                "image/svg+xml" => ".svg",
                "image/webp" => ".webp",
                _ => ".png"
            };
        }

        private static void AddImageFallbackText(WordParagraph paragraph, string? alt, string fallback, MarkdownToWordOptions options) {
            var fallbackText = !string.IsNullOrEmpty(alt) ? alt! : fallback ?? string.Empty;
            var text = paragraph.AddText(fallbackText);
            var defaultFont = ResolveDefaultFontFamily(options);
            if (!string.IsNullOrEmpty(defaultFont)) {
                text.SetFontFamily(defaultFont!);
            }
        }

        private static double? ResolveContextWidthLimitPixels(
            MarkdownImageLayoutOptions layout,
            double pageContentWidthPixels,
            int listLevel,
            int quoteDepth) {
            if (layout.FitMode == MarkdownImageFitMode.None || pageContentWidthPixels <= 0) {
                return null;
            }

            if (layout.FitMode == MarkdownImageFitMode.PageContentWidth) {
                return pageContentWidthPixels;
            }

            var levels = Math.Max(0, listLevel) + Math.Max(0, quoteDepth);
            if (levels == 0) {
                return pageContentWidthPixels;
            }

            var indentPixels = levels * (IndentTwipsPerLevel * PixelsPerInch / TwipsPerInch);
            return Math.Max(MinimumContentWidthPixels, pageContentWidthPixels - indentPixels);
        }

        private static bool TryGetImageDimensionsFromFile(string filePath, out double width, out double height) {
            width = 0;
            height = 0;
            if (OfficeImageReader.TryIdentify(File.ReadAllBytes(filePath), filePath, out var image)) {
                width = image.Width;
                height = image.Height;
                return width > 0 && height > 0;
            }

            return false;
        }

        private static bool TryGetImageDimensionsFromBytes(byte[] data, out double width, out double height) {
            width = 0;
            height = 0;
            if (OfficeImageReader.TryIdentify(data, null, out var image)) {
                width = image.Width;
                height = image.Height;
                return width > 0 && height > 0;
            }

            return false;
        }

        private static bool NormalizePositiveDimension(double? value, out double normalized) {
            normalized = 0;
            if (!value.HasValue || double.IsNaN(value.Value) || double.IsInfinity(value.Value)) {
                return false;
            }

            if (value.Value <= 0) {
                return false;
            }

            normalized = value.Value;
            return true;
        }

        private static void ResolveImageDimensions(
            MarkdownToWordOptions options,
            string source,
            string context,
            double? requestedWidth,
            double? requestedHeight,
            double? naturalWidth,
            double? naturalHeight,
            double pageContentWidthPixels,
            double? contextWidthLimitPixels,
            out double? finalWidth,
            out double? finalHeight,
            out bool scaledByLayout) {
            var layout = options.ImageLayout ?? new MarkdownImageLayoutOptions();
            finalWidth = null;
            finalHeight = null;
            scaledByLayout = false;

            var hasNaturalWidth = NormalizePositiveDimension(naturalWidth, out var naturalWidthPx);
            var hasNaturalHeight = NormalizePositiveDimension(naturalHeight, out var naturalHeightPx);
            var hasRequestedWidth = NormalizePositiveDimension(requestedWidth, out var requestedWidthPx);
            var hasRequestedHeight = NormalizePositiveDimension(requestedHeight, out var requestedHeightPx);

            if (layout.HintPrecedence == MarkdownImageHintPrecedence.LayoutThenMarkdown) {
                if (hasNaturalWidth) {
                    finalWidth = naturalWidthPx;
                }
                if (hasNaturalHeight) {
                    finalHeight = naturalHeightPx;
                }

                if (hasRequestedWidth) {
                    finalWidth = requestedWidthPx;
                }
                if (hasRequestedHeight) {
                    finalHeight = requestedHeightPx;
                }
            } else {
                if (hasRequestedWidth) {
                    finalWidth = requestedWidthPx;
                } else if (hasNaturalWidth) {
                    finalWidth = naturalWidthPx;
                }

                if (hasRequestedHeight) {
                    finalHeight = requestedHeightPx;
                } else if (hasNaturalHeight) {
                    finalHeight = naturalHeightPx;
                }
            }

            if (finalWidth.HasValue && !finalHeight.HasValue && hasNaturalWidth && hasNaturalHeight) {
                finalHeight = naturalHeightPx * (finalWidth.Value / naturalWidthPx);
            } else if (!finalWidth.HasValue && finalHeight.HasValue && hasNaturalWidth && hasNaturalHeight) {
                finalWidth = naturalWidthPx * (finalHeight.Value / naturalHeightPx);
            }

            double? effectiveMaxWidth = null;
            double? effectiveMaxHeight = null;

            if (NormalizePositiveDimension(layout.MaxWidthPixels, out var maxWidth)) {
                effectiveMaxWidth = maxWidth;
            }
            if (NormalizePositiveDimension(layout.MaxHeightPixels, out var maxHeight)) {
                effectiveMaxHeight = maxHeight;
            }
            if (NormalizePositiveDimension(layout.MaxWidthPercentOfContent, out var maxWidthPercent)) {
                var widthBaseline = NormalizePositiveDimension(contextWidthLimitPixels, out var contextWidth)
                    ? contextWidth
                    : (pageContentWidthPixels > 0 ? pageContentWidthPixels : 0);
                if (widthBaseline > 0) {
                    var percentCapWidth = widthBaseline * (maxWidthPercent / 100d);
                    if (percentCapWidth > 0) {
                        effectiveMaxWidth = effectiveMaxWidth.HasValue
                            ? Math.Min(effectiveMaxWidth.Value, percentCapWidth)
                            : percentCapWidth;
                    }
                }
            }
            if (NormalizePositiveDimension(contextWidthLimitPixels, out var contextMaxWidth)) {
                effectiveMaxWidth = effectiveMaxWidth.HasValue ? Math.Min(effectiveMaxWidth.Value, contextMaxWidth) : contextMaxWidth;
            }

            if (!layout.AllowUpscale) {
                if (hasNaturalWidth) {
                    effectiveMaxWidth = effectiveMaxWidth.HasValue ? Math.Min(effectiveMaxWidth.Value, naturalWidthPx) : naturalWidthPx;
                }
                if (hasNaturalHeight) {
                    effectiveMaxHeight = effectiveMaxHeight.HasValue ? Math.Min(effectiveMaxHeight.Value, naturalHeightPx) : naturalHeightPx;
                }
            }

            if (finalWidth.HasValue && finalHeight.HasValue) {
                var scale = 1d;
                if (effectiveMaxWidth.HasValue && finalWidth.Value > effectiveMaxWidth.Value) {
                    scale = Math.Min(scale, effectiveMaxWidth.Value / finalWidth.Value);
                }
                if (effectiveMaxHeight.HasValue && finalHeight.Value > effectiveMaxHeight.Value) {
                    scale = Math.Min(scale, effectiveMaxHeight.Value / finalHeight.Value);
                }
                if (scale < 1d) {
                    finalWidth *= scale;
                    finalHeight *= scale;
                    scaledByLayout = true;
                }
            } else {
                if (finalWidth.HasValue && effectiveMaxWidth.HasValue && finalWidth.Value > effectiveMaxWidth.Value) {
                    finalWidth = effectiveMaxWidth.Value;
                    scaledByLayout = true;
                }
                if (finalHeight.HasValue && effectiveMaxHeight.HasValue && finalHeight.Value > effectiveMaxHeight.Value) {
                    finalHeight = effectiveMaxHeight.Value;
                    scaledByLayout = true;
                }
            }

            if (finalWidth.HasValue && finalWidth.Value <= 0) {
                finalWidth = null;
            }
            if (finalHeight.HasValue && finalHeight.Value <= 0) {
                finalHeight = null;
            }

            if (options.OnImageLayoutDiagnostic != null) {
                options.OnImageLayoutDiagnostic(new MarkdownImageLayoutDiagnostic {
                    Source = source,
                    Context = context,
                    RequestedWidthPixels = hasRequestedWidth ? requestedWidthPx : null,
                    RequestedHeightPixels = hasRequestedHeight ? requestedHeightPx : null,
                    NaturalWidthPixels = hasNaturalWidth ? naturalWidthPx : null,
                    NaturalHeightPixels = hasNaturalHeight ? naturalHeightPx : null,
                    EffectiveMaxWidthPixels = effectiveMaxWidth,
                    EffectiveMaxHeightPixels = effectiveMaxHeight,
                    FinalWidthPixels = finalWidth,
                    FinalHeightPixels = finalHeight,
                    ScaledByLayout = scaledByLayout
                });
            }
        }
    }
}
