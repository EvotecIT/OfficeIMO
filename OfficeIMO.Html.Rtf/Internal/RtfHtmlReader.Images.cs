namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void AddImage(IElement token) {
            string source = HtmlImageSourceResolver.ResolveImageSource(token, _baseUri, _options.GetResourceUrlPolicy());
            if (string.IsNullOrWhiteSpace(source) || !TryReadDataImage(source, out RtfImageFormat format, out byte[]? data)) {
                string? alt = GetAttribute(token, "alt");
                if (!string.IsNullOrWhiteSpace(alt)) {
                    AppendText(alt!);
                }

                return;
            }

            RtfImage image = EnsureInlineParagraph().AddImage(format, data!);
            image.Description = GetAttribute(token, "alt");
            ApplyImageSize(token, image);
        }

        private static void ApplyImageSize(IElement token, RtfImage image) {
            string? width = GetAttribute(token, "width");
            if (!string.IsNullOrWhiteSpace(width) && HtmlStyleDeclarationParser.TryParseTwips(width!, out int widthTwips)) {
                image.DesiredWidthTwips = widthTwips;
                if (TryParsePositiveInteger(width!, out int sourceWidth)) {
                    image.SourceWidth = sourceWidth;
                }
            }

            string? height = GetAttribute(token, "height");
            if (!string.IsNullOrWhiteSpace(height) && HtmlStyleDeclarationParser.TryParseTwips(height!, out int heightTwips)) {
                image.DesiredHeightTwips = heightTwips;
                if (TryParsePositiveInteger(height!, out int sourceHeight)) {
                    image.SourceHeight = sourceHeight;
                }
            }

            HtmlStyleDeclaration style = HtmlStyleDeclarationParser.Parse(GetAttribute(token, "style"));
            if (style.TableWidth.HasValue && style.TableWidthUnit == RtfTableWidthUnit.Twips) {
                image.DesiredWidthTwips = style.TableWidth.Value;
            }

            if (style.TableHeightTwips.HasValue) {
                image.DesiredHeightTwips = style.TableHeightTwips.Value;
            }
        }

        private static bool TryParsePositiveInteger(string value, out int result) {
            string normalized = value.Trim();
            result = 0;
            if (normalized.IndexOfAny(new[] { '.', ',', '%', ' ', '\t', '\r', '\n' }) >= 0 ||
                !int.TryParse(normalized, out int parsed) ||
                parsed <= 0) {
                return false;
            }

            result = parsed;
            return true;
        }

        private static bool TryReadDataImage(string source, out RtfImageFormat format, out byte[]? data) {
            format = RtfImageFormat.Unknown;
            data = null;
            if (!HtmlImageDataUri.TryParse(source, out HtmlImageDataUri dataUri) || !dataUri.IsBase64) {
                return false;
            }

            if (dataUri.MediaType.Equals("image/png", StringComparison.OrdinalIgnoreCase)) {
                format = RtfImageFormat.Png;
            } else if (dataUri.MediaType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) || dataUri.MediaType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)) {
                format = RtfImageFormat.Jpeg;
            } else {
                return false;
            }

            if (dataUri.TryDecodeBytes(out data)) {
                return true;
            }

            if (data == null || data.Length == 0) {
                format = RtfImageFormat.Unknown;
                data = null;
            }

            return false;
        }
    }
}
