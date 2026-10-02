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
            var metadata = RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-picture"));
            if (metadata.TryGetValue("version", out string? version) && version == "1") {
                image.SourceWidth = ReadInt(metadata, "SourceWidth");
                image.SourceHeight = ReadInt(metadata, "SourceHeight");
                image.DesiredWidthTwips = ReadInt(metadata, "DesiredWidthTwips");
                image.DesiredHeightTwips = ReadInt(metadata, "DesiredHeightTwips");
                image.ScaleXPercent = ReadInt(metadata, "ScaleXPercent");
                image.ScaleYPercent = ReadInt(metadata, "ScaleYPercent");
                image.CropLeftTwips = ReadInt(metadata, "CropLeftTwips");
                image.CropTopTwips = ReadInt(metadata, "CropTopTwips");
                image.CropRightTwips = ReadInt(metadata, "CropRightTwips");
                image.CropBottomTwips = ReadInt(metadata, "CropBottomTwips");
            }
        }

        private void ApplyImageSize(IElement token, RtfImage image) {
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

            if (image.DesiredWidthTwips.HasValue || image.DesiredHeightTwips.HasValue ||
                !OfficeImageReader.TryIdentifyByContent(image.Data, null, out OfficeImageInfo info) ||
                info.Width <= 0 || info.Height <= 0) return;

            // Resolve the actual section/document text area, using the RTF defaults
            // also materialized by RtfDocument.Merge.Sections for omitted page setup.
            RtfPageSetup page = _document.PageSetup;
            RtfPageSetup? section = _currentSection?.PageSetup;
            double maxTextWidthTwips = Math.Max(1D,
                (section?.PaperWidthTwips ?? page.PaperWidthTwips ?? 12240D)
                - (section?.MarginLeftTwips ?? page.MarginLeftTwips ?? 1800D)
                - (section?.MarginRightTwips ?? page.MarginRightTwips ?? 1800D)
                - (section?.GutterWidthTwips ?? page.GutterWidthTwips ?? 0D));
            double maxTextHeightTwips = Math.Max(1D,
                (section?.PaperHeightTwips ?? page.PaperHeightTwips ?? 15840D)
                - (section?.MarginTopTwips ?? page.MarginTopTwips ?? 1440D)
                - (section?.MarginBottomTwips ?? page.MarginBottomTwips ?? 1440D));
            double naturalWidth = info.Width * 1440D / info.DpiX;
            double naturalHeight = info.Height * 1440D / info.DpiY;
            double scale = Math.Min(1D,
                Math.Min(maxTextWidthTwips / naturalWidth, maxTextHeightTwips / naturalHeight));
            if (scale >= 1D) return;

            image.SourceWidth = info.Width;
            image.SourceHeight = info.Height;
            image.DesiredWidthTwips = Math.Max(1, (int)Math.Round(naturalWidth * scale));
            image.DesiredHeightTwips = Math.Max(1, (int)Math.Round(naturalHeight * scale));
            _options.AddDiagnostic("HtmlRtfImageFittedToPage",
                "An unstyled HTML image was proportionally fitted to the RTF page text area.",
                HtmlRenderStyleResolver.DescribeSource(token), action: RtfConversionAction.Substituted);
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
