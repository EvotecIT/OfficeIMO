using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyWorksheetHeaderFooter(PdfCore.PdfPageBuilder page, ExcelSheet.ExcelHeaderFooterInfo? headerFooter, string sheetName, string? workbookPath, ExcelPdfExportOptions options, ExcelSheetPageSetup? pageSetup = null) {
            if (headerFooter == null) {
                return;
            }

            PreparedHeaderFooterImages preparedImages = PrepareHeaderFooterImages(headerFooter, sheetName, options);
            PdfCore.PdfPageMargins pageMargins = GetEffectiveMargins(options, pageSetup);
            // Header pictures stand inline on the header line, which starts at the header margin,
            // so the room they have is the band between the header margin and the sheet.
            double headerMarginPoints = GetHeaderFooterMarginPoints(pageSetup, header: true);
            // Keep the line's descent and a small gap clear of the sheet below it.
            double headerLineClearance = (options.RunningLineFontSize ?? options.PdfOptions?.HeaderFontSize ?? DefaultHeaderFooterFontSize) * 0.25D + 3D;
            FitHeaderFooterImagesToMargin(preparedImages, header: true, Math.Max(0D, pageMargins.Top - headerMarginPoints - headerLineClearance) + HeaderTextAllowance, sheetName, options);
            double tallestHeaderImage = new[] { preparedImages.HeaderLeft, preparedImages.HeaderCenter, preparedImages.HeaderRight }
                .Where(image => image != null).Select(image => image!.HeightPoints).DefaultIfEmpty(0D).Max();
            double? footerImageOffset = FitHeaderFooterImagesToMargin(preparedImages, header: false, pageMargins.Bottom, sheetName, options);

            HeaderFooterZones? headerZones = options.UseWorksheetHeadersAndFooters ? ConvertHeaderFooterZones(headerFooter.HeaderLeft, headerFooter.HeaderCenter, headerFooter.HeaderRight, sheetName, workbookPath, options, "header") : null;
            var firstHeaderZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage
                ? ConvertHeaderFooterZones(headerFooter.FirstHeaderLeft, headerFooter.FirstHeaderCenter, headerFooter.FirstHeaderRight, sheetName, workbookPath, options, "first-page header")
                : null;
            var evenHeaderZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven
                ? ConvertHeaderFooterZones(headerFooter.EvenHeaderLeft, headerFooter.EvenHeaderCenter, headerFooter.EvenHeaderRight, sheetName, workbookPath, options, "even-page header")
                : null;
            bool hasFirstHeaderVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage;
            bool hasEvenHeaderVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven;
            if (HasAnyText(headerZones) || hasFirstHeaderVariant || hasEvenHeaderVariant || preparedImages.HasHeaderImages) {
                HeaderFooterLineStyle? headerStyle = WithBodyFont(ResolveSharedHeaderFooterStyle(new[] { headerZones, firstHeaderZones, evenHeaderZones }, sheetName, options, "header"), options);
                double headerOffset = ResolveHeaderLineOffset(pageMargins.Top, headerMarginPoints, headerStyle?.FontSize ?? options.PdfOptions?.HeaderFontSize, tallestHeaderImage);
                page.Header(header => {
                    ApplyHeaderFooterStyle(header, headerStyle);
                    header.Offset(headerOffset);

                    if (HasAnyText(headerZones)) {
                        header.Zones(headerZones!.Left, headerZones.Center, headerZones.Right);
                    }

                    if (HasAnyText(firstHeaderZones)) {
                        header.FirstPageZones(firstHeaderZones!.Left, firstHeaderZones.Center, firstHeaderZones.Right);
                    } else if (hasFirstHeaderVariant) {
                        header.FirstPageText(string.Empty);
                    }

                    if (HasAnyText(evenHeaderZones)) {
                        header.EvenPagesZones(evenHeaderZones!.Left, evenHeaderZones.Center, evenHeaderZones.Right);
                    } else if (hasEvenHeaderVariant) {
                        header.EvenPagesText(string.Empty);
                    }

                    AddHeaderImage(header, preparedImages.HeaderLeft, PdfCore.PdfAlign.Left, headerFooter.HeaderLeft, headerFooter.FirstHeaderLeft, headerFooter.EvenHeaderLeft);
                    AddHeaderImage(header, preparedImages.HeaderCenter, PdfCore.PdfAlign.Center, headerFooter.HeaderCenter, headerFooter.FirstHeaderCenter, headerFooter.EvenHeaderCenter);
                    AddHeaderImage(header, preparedImages.HeaderRight, PdfCore.PdfAlign.Right, headerFooter.HeaderRight, headerFooter.FirstHeaderRight, headerFooter.EvenHeaderRight);
                });
            }

            HeaderFooterZones? footerZones = options.UseWorksheetHeadersAndFooters ? ConvertHeaderFooterZones(headerFooter.FooterLeft, headerFooter.FooterCenter, headerFooter.FooterRight, sheetName, workbookPath, options, "footer") : null;
            var firstFooterZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage
                ? ConvertHeaderFooterZones(headerFooter.FirstFooterLeft, headerFooter.FirstFooterCenter, headerFooter.FirstFooterRight, sheetName, workbookPath, options, "first-page footer")
                : null;
            var evenFooterZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven
                ? ConvertHeaderFooterZones(headerFooter.EvenFooterLeft, headerFooter.EvenFooterCenter, headerFooter.EvenFooterRight, sheetName, workbookPath, options, "even-page footer")
                : null;
            bool hasFirstFooterVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage;
            bool hasEvenFooterVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven;
            if (HasAnyText(footerZones) || hasFirstFooterVariant || hasEvenFooterVariant || preparedImages.HasFooterImages) {
                HeaderFooterLineStyle? footerStyle = WithBodyFont(ResolveSharedHeaderFooterStyle(new[] { footerZones, firstFooterZones, evenFooterZones }, sheetName, options, "footer"), options);
                double footerOffset = ResolveHeaderFooterLineOffset(header: false, pageMargins.Bottom, GetHeaderFooterMarginPoints(pageSetup, header: false), footerStyle?.FontSize ?? options.PdfOptions?.FooterFontSize, footerImageOffset);
                page.Footer(footer => {
                    ApplyHeaderFooterStyle(footer, footerStyle);
                    footer.Offset(footerOffset);

                    if (HasAnyText(footerZones)) {
                        footer.Zones(footerZones!.Left, footerZones.Center, footerZones.Right);
                    }

                    if (HasAnyText(firstFooterZones)) {
                        footer.FirstPageZones(firstFooterZones!.Left, firstFooterZones.Center, firstFooterZones.Right);
                    } else if (hasFirstFooterVariant) {
                        footer.FirstPageText(string.Empty);
                    }

                    if (HasAnyText(evenFooterZones)) {
                        footer.EvenPagesZones(evenFooterZones!.Left, evenFooterZones.Center, evenFooterZones.Right);
                    } else if (hasEvenFooterVariant) {
                        footer.EvenPagesText(string.Empty);
                    }

                    AddFooterImage(footer, preparedImages.FooterLeft, PdfCore.PdfAlign.Left, headerFooter.FooterLeft, headerFooter.FirstFooterLeft, headerFooter.EvenFooterLeft);
                    AddFooterImage(footer, preparedImages.FooterCenter, PdfCore.PdfAlign.Center, headerFooter.FooterCenter, headerFooter.FirstFooterCenter, headerFooter.EvenFooterCenter);
                    AddFooterImage(footer, preparedImages.FooterRight, PdfCore.PdfAlign.Right, headerFooter.FooterRight, headerFooter.FirstFooterRight, headerFooter.EvenFooterRight);
                });
            }
        }

        private static void AddHeaderImage(PdfCore.PdfHeaderBuilder header, PreparedHeaderFooterImage? image, PdfCore.PdfAlign align, string defaultZone, string firstZone, string evenZone) {
            if (image != null) {
                bool defaultHasImage = HasPicturePlaceholder(defaultZone);
                bool firstHasImage = HasPicturePlaceholder(firstZone);
                bool evenHasImage = HasPicturePlaceholder(evenZone);
                bool hasSpecificPlaceholder = defaultHasImage || firstHasImage || evenHasImage;
                if (!hasSpecificPlaceholder || defaultHasImage) {
                    header.Image(CreateInlineHeaderImage(image, align));
                }

                if (firstHasImage) {
                    header.FirstPageImage(CreateInlineHeaderImage(image, align));
                }

                if (evenHasImage) {
                    header.EvenPagesImage(CreateInlineHeaderImage(image, align));
                }
            }
        }

        // Excel sets a header picture inline with the section's text, standing on its baseline.
        private static PdfCore.PdfHeaderFooterImage CreateInlineHeaderImage(PreparedHeaderFooterImage image, PdfCore.PdfAlign align) =>
            new PdfCore.PdfHeaderFooterImage(image.Bytes, image.WidthPoints, image.HeightPoints, align) { RestsOnHeaderBaseline = true };

        private static void AddFooterImage(PdfCore.PdfFooterBuilder footer, PreparedHeaderFooterImage? image, PdfCore.PdfAlign align, string defaultZone, string firstZone, string evenZone) {
            if (image != null) {
                bool defaultHasImage = HasPicturePlaceholder(defaultZone);
                bool firstHasImage = HasPicturePlaceholder(firstZone);
                bool evenHasImage = HasPicturePlaceholder(evenZone);
                bool hasSpecificPlaceholder = defaultHasImage || firstHasImage || evenHasImage;
                if (!hasSpecificPlaceholder || defaultHasImage) {
                    footer.Image(image.Bytes, image.WidthPoints, image.HeightPoints, align);
                }

                if (firstHasImage) {
                    footer.FirstPageImage(image.Bytes, image.WidthPoints, image.HeightPoints, align);
                }

                if (evenHasImage) {
                    footer.EvenPagesImage(image.Bytes, image.WidthPoints, image.HeightPoints, align);
                }
            }
        }

        private static bool HasPicturePlaceholder(string? text) =>
            text?.IndexOf("&G", StringComparison.Ordinal) >= 0;

        private static PreparedHeaderFooterImages PrepareHeaderFooterImages(
            ExcelSheet.ExcelHeaderFooterInfo headerFooter,
            string sheetName,
            ExcelPdfExportOptions options) {
            if (!options.UseWorksheetHeaderFooterImages) return new PreparedHeaderFooterImages();

            return new PreparedHeaderFooterImages {
                HeaderLeft = PrepareHeaderFooterImage(headerFooter.HeaderLeftImage, sheetName, "header left", options),
                HeaderCenter = PrepareHeaderFooterImage(headerFooter.HeaderCenterImage, sheetName, "header center", options),
                HeaderRight = PrepareHeaderFooterImage(headerFooter.HeaderRightImage, sheetName, "header right", options),
                FooterLeft = PrepareHeaderFooterImage(headerFooter.FooterLeftImage, sheetName, "footer left", options),
                FooterCenter = PrepareHeaderFooterImage(headerFooter.FooterCenterImage, sheetName, "footer center", options),
                FooterRight = PrepareHeaderFooterImage(headerFooter.FooterRightImage, sheetName, "footer right", options)
            };
        }

        private static PreparedHeaderFooterImage? PrepareHeaderFooterImage(
            ExcelSheet.ExcelHeaderFooterImageInfo? image,
            string sheetName,
            string location,
            ExcelPdfExportOptions options) {
            if (image == null) return null;

            if (image.Bytes.Length > 0 &&
                image.WidthPoints > 0D &&
                image.HeightPoints > 0D &&
                IsPdfSupportedImageContentType(image.ContentType) &&
                TryPreparePdfImageBytes(image.Bytes, image.ContentType, out byte[] preparedBytes, out _)) {
                return new PreparedHeaderFooterImage(preparedBytes, image.WidthPoints, image.HeightPoints);
            }

            AddWarning(
                options,
                sheetName,
                "worksheet-header-footer-image",
                $"The {location} image was not exported because it is not a supported PDF image payload. ContentType='{image.ContentType}', WidthPoints={image.WidthPoints.ToString(CultureInfo.InvariantCulture)}, HeightPoints={image.HeightPoints.ToString(CultureInfo.InvariantCulture)}, Bytes={image.Bytes.Length.ToString(CultureInfo.InvariantCulture)}.");
            return null;
        }

        private static bool HasAnyText(params string?[] values) {
            foreach (string? value in values) {
                if (!string.IsNullOrWhiteSpace(value)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasAnyText(HeaderFooterZones? zones) {
            if (zones == null) {
                return false;
            }

            return HasAnyText(zones.Left, zones.Center, zones.Right);
        }


        private static HeaderFooterZones ConvertHeaderFooterZones(string? left, string? center, string? right, string sheetName, string? workbookPath, ExcelPdfExportOptions options, string scope) {
            HeaderFooterZone leftZone = ConvertHeaderFooterText(left, sheetName, workbookPath, options, scope, "left");
            HeaderFooterZone centerZone = ConvertHeaderFooterText(center, sheetName, workbookPath, options, scope, "center");
            HeaderFooterZone rightZone = ConvertHeaderFooterText(right, sheetName, workbookPath, options, scope, "right");
            return new HeaderFooterZones(
                leftZone.Text,
                centerZone.Text,
                rightZone.Text,
                ResolveSharedHeaderFooterZoneStyle(new[] { leftZone, centerZone, rightZone }, sheetName, options, scope));
        }

        private static HeaderFooterZone ConvertHeaderFooterText(string? text, string sheetName, string? workbookPath, ExcelPdfExportOptions options, string scope, string zone) {
            if (string.IsNullOrWhiteSpace(text)) {
                return HeaderFooterZone.Empty;
            }

            var builder = new System.Text.StringBuilder(text!.Length);
            var style = new HeaderFooterLineStyle();
            bool unsupportedFormatting = false;
            bool canApplyLineStyle = true;
            bool hasVisibleContent = false;
            DateTime? headerFooterDateTime = null;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch != '&' || i + 1 >= text.Length) {
                    builder.Append(ch);
                    if (!char.IsWhiteSpace(ch)) {
                        hasVisibleContent = true;
                    }

                    continue;
                }

                char token = text[++i];
                switch (token) {
                    case '&':
                        if (i + 1 < text.Length && char.IsDigit(text[i + 1])) {
                            if (TryReadHeaderFooterFontSize(text, i + 1, out double fontSize, out int fontSizeEndIndex)) {
                                ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.FontSize = fontSize);
                                i = fontSizeEndIndex;
                            } else {
                                unsupportedFormatting = true;
                            }
                        } else {
                            builder.Append('&');
                            hasVisibleContent = true;
                        }

                        break;
                    case 'P':
                        builder.Append("{page}");
                        hasVisibleContent = true;
                        break;
                    case 'N':
                        builder.Append("{pages}");
                        hasVisibleContent = true;
                        break;
                    case 'A':
                        builder.Append(NormalizeHeaderFooterFieldText(sheetName));
                        hasVisibleContent = true;
                        break;
                    case 'D':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterDateTime(options, ref headerFooterDateTime).ToString("d", CultureInfo.CurrentCulture)));
                        hasVisibleContent = true;
                        break;
                    case 'T':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterDateTime(options, ref headerFooterDateTime).ToString("t", CultureInfo.CurrentCulture)));
                        hasVisibleContent = true;
                        break;
                    case 'F':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterFileName(workbookPath)));
                        hasVisibleContent = true;
                        break;
                    case 'Z':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterDirectory(workbookPath)));
                        hasVisibleContent = true;
                        break;
                    case 'G':
                        break;
                    case 'B':
                        ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.Bold = !s.Bold);
                        break;
                    case 'I':
                        ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.Italic = !s.Italic);
                        break;
                    case 'U':
                    case 'S':
                        unsupportedFormatting = true;
                        if (hasVisibleContent) {
                            canApplyLineStyle = false;
                        }

                        break;
                    case 'K':
                        if (TryReadHeaderFooterColor(text, i + 1, out PdfCore.PdfColor color, out int colorEndIndex)) {
                            ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.Color = color);
                            i = colorEndIndex;
                        } else {
                            unsupportedFormatting = true;
                            i = SkipExcelHeaderFooterColorToken(text, i);
                        }

                        break;
                    case '"':
                        if (TryReadHeaderFooterQuotedToken(text, i, out string quotedToken, out int quotedEndIndex) &&
                            TryApplyHeaderFooterFontToken(style, quotedToken)) {
                            ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, _ => { });
                            i = quotedEndIndex;
                        } else {
                            unsupportedFormatting = true;
                            i = SkipExcelHeaderFooterQuotedToken(text, i);
                        }

                        break;
                    default:
                        if (char.IsDigit(token)) {
                            if (TryReadHeaderFooterFontSize(text, i, out double fontSize, out int fontSizeEndIndex)) {
                                ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.FontSize = fontSize);
                                i = fontSizeEndIndex;
                            } else {
                                unsupportedFormatting = true;
                            }
                        } else {
                            builder.Append(token);
                            hasVisibleContent = true;
                        }
                        break;
                }
            }

            if (unsupportedFormatting) {
                AddWarning(
                    options,
                    sheetName,
                    "worksheet-header-footer-formatting",
                    $"Excel header/footer formatting in the {scope} {zone} zone was simplified. Text, page tokens, total-page tokens, sheet-name, date/time, and workbook file fields are preserved, but rich formatting is not exported yet.");
            }

            string result = builder.ToString().Trim();
            return result.Length == 0
                ? HeaderFooterZone.Empty
                : new HeaderFooterZone(result, canApplyLineStyle && style.HasAnyStyle ? style : null);
        }

        private static void ApplyHeaderFooterStyleToken(HeaderFooterLineStyle style, bool hasVisibleContent, ref bool unsupportedFormatting, ref bool canApplyLineStyle, Action<HeaderFooterLineStyle> apply) {
            if (hasVisibleContent) {
                unsupportedFormatting = true;
                canApplyLineStyle = false;
                return;
            }

            apply(style);
        }

        private static HeaderFooterLineStyle? ResolveSharedHeaderFooterZoneStyle(HeaderFooterZone[] zones, string sheetName, ExcelPdfExportOptions options, string scope) {
            HeaderFooterLineStyle? shared = null;
            bool hasStyle = false;
            bool hasUnstyledText = false;
            foreach (HeaderFooterZone zone in zones) {
                if (string.IsNullOrWhiteSpace(zone.Text)) {
                    continue;
                }

                if (zone.Style == null) {
                    hasUnstyledText = true;
                    continue;
                }

                if (!hasStyle) {
                    shared = zone.Style;
                    hasStyle = true;
                    continue;
                }

                if (!HeaderFooterLineStyle.Equals(shared, zone.Style)) {
                    AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                    return null;
                }
            }

            if (hasStyle && hasUnstyledText) {
                AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                return null;
            }

            return shared;
        }

        private static HeaderFooterLineStyle? ResolveSharedHeaderFooterStyle(HeaderFooterZones?[] zoneSets, string sheetName, ExcelPdfExportOptions options, string scope) {
            HeaderFooterLineStyle? shared = null;
            bool hasStyle = false;
            bool hasUnstyledText = false;
            foreach (HeaderFooterZones? zones in zoneSets) {
                if (!HasAnyText(zones)) {
                    continue;
                }

                if (zones!.Style == null) {
                    hasUnstyledText = true;
                    continue;
                }

                if (!hasStyle) {
                    shared = zones.Style;
                    hasStyle = true;
                    continue;
                }

                if (!HeaderFooterLineStyle.Equals(shared, zones.Style)) {
                    AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                    return null;
                }
            }

            if (hasStyle && hasUnstyledText) {
                AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                return null;
            }

            return shared;
        }

        private static void ApplyHeaderFooterStyle(PdfCore.PdfHeaderBuilder header, HeaderFooterLineStyle? style) {
            if (style == null) {
                return;
            }

            if (style.FontSize.HasValue) {
                header.FontSize(style.FontSize.Value);
            }

            if (style.Color.HasValue) {
                header.Color(style.Color.Value);
            }

            if (style.Font.HasValue) {
                header.Font(style.Font.Value);
            }

            if (!string.IsNullOrWhiteSpace(style.FontFamilyName)) {
                header.FontFamily(style.FontFamilyName!);
            }
        }

        private static void ApplyHeaderFooterStyle(PdfCore.PdfFooterBuilder footer, HeaderFooterLineStyle? style) {
            if (style == null) {
                return;
            }

            if (style.FontSize.HasValue) {
                footer.FontSize(style.FontSize.Value);
            }

            if (style.Color.HasValue) {
                footer.Color(style.Color.Value);
            }

            if (style.Font.HasValue) {
                footer.Font(style.Font.Value);
            }

            if (!string.IsNullOrWhiteSpace(style.FontFamilyName)) {
                footer.FontFamily(style.FontFamilyName!);
            }
        }

        private static void AddMixedHeaderFooterFormattingWarning(ExcelPdfExportOptions options, string sheetName, string scope) {
            AddWarning(
                options,
                sheetName,
                "worksheet-header-footer-formatting",
                $"Excel header/footer formatting in the {scope} uses mixed or partial styles that cannot be represented as one PDF header/footer line style yet. Text is preserved, but rich formatting is simplified.");
        }

        private static DateTime GetHeaderFooterDateTime(ExcelPdfExportOptions options, ref DateTime? dateTime) {
            if (!dateTime.HasValue) {
                dateTime = options.HeaderFooterDateTimeProvider != null
                    ? options.HeaderFooterDateTimeProvider()
                    : DateTime.Now;
            }

            return dateTime.Value;
        }

        private static string GetHeaderFooterFileName(string? workbookPath) {
            if (string.IsNullOrWhiteSpace(workbookPath)) {
                return "Workbook";
            }

            string fileName = Path.GetFileName(workbookPath!);
            return string.IsNullOrWhiteSpace(fileName) ? "Workbook" : fileName;
        }

        private static string GetHeaderFooterDirectory(string? workbookPath) {
            if (string.IsNullOrWhiteSpace(workbookPath)) {
                return string.Empty;
            }

            string? directory = Path.GetDirectoryName(Path.GetFullPath(workbookPath!));
            return directory ?? string.Empty;
        }

        private static string NormalizeHeaderFooterFieldText(string text) =>
            text.Replace('\u00A0', ' ').Replace('\u202F', ' ');

        private static bool TryReadHeaderFooterFontSize(string text, int startIndex, out double fontSize, out int endIndex) {
            fontSize = 0D;
            endIndex = startIndex;
            int index = startIndex;
            while (index < text.Length && char.IsDigit(text[index])) {
                index++;
            }

            if (index == startIndex) {
                return false;
            }

            string token = text.Substring(startIndex, index - startIndex);
            endIndex = index - 1;
            return double.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out fontSize) && fontSize > 0D;
        }

        private static bool TryReadHeaderFooterColor(string text, int startIndex, out PdfCore.PdfColor color, out int endIndex) {
            color = default;
            endIndex = startIndex;
            if (startIndex + 6 > text.Length) {
                return false;
            }

            for (int i = 0; i < 6; i++) {
                if (!IsHexDigit(text[startIndex + i])) {
                    return false;
                }
            }

            byte red = byte.Parse(text.Substring(startIndex, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte green = byte.Parse(text.Substring(startIndex + 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte blue = byte.Parse(text.Substring(startIndex + 4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            color = PdfCore.PdfColor.FromRgb(red, green, blue);
            endIndex = startIndex + 5;
            return true;
        }

        private static bool TryReadHeaderFooterQuotedToken(string text, int quoteIndex, out string value, out int endIndex) {
            value = string.Empty;
            endIndex = quoteIndex;
            int closingIndex = quoteIndex + 1;
            while (closingIndex < text.Length && text[closingIndex] != '"') {
                closingIndex++;
            }

            if (closingIndex >= text.Length) {
                return false;
            }

            value = text.Substring(quoteIndex + 1, closingIndex - quoteIndex - 1);
            endIndex = closingIndex;
            return true;
        }

        private static bool TryApplyHeaderFooterFontToken(HeaderFooterLineStyle style, string token) {
            if (string.IsNullOrWhiteSpace(token)) {
                return false;
            }

            string[] parts = token.Split(new[] { ',' }, 2);
            string familyName = parts[0].Trim();
            if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont family)) {
                style.FontFamily = family;
            }

            style.FontFamilyName = familyName;
            if (parts.Length > 1) {
                string fontStyle = parts[1];
                style.Bold = fontStyle.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0;
                style.Italic = fontStyle.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               fontStyle.IndexOf("oblique", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            return true;
        }

        private static int SkipExcelHeaderFooterColorToken(string text, int index) {
            int skipped = 0;
            while (index + 1 < text.Length && skipped < 6 && IsHexDigit(text[index + 1])) {
                index++;
                skipped++;
            }

            return index;
        }

        private static int SkipExcelHeaderFooterQuotedToken(string text, int index) {
            while (index + 1 < text.Length) {
                index++;
                if (text[index] == '"') {
                    break;
                }
            }

            return index;
        }

        private static int SkipExcelHeaderFooterFontSizeToken(string text, int index) {
            while (index + 1 < text.Length && char.IsDigit(text[index + 1])) {
                index++;
            }

            return index;
        }

        private static bool IsHexDigit(char value) {
            return (value >= '0' && value <= '9') ||
                   (value >= 'a' && value <= 'f') ||
                   (value >= 'A' && value <= 'F');
        }

        /// <summary>Gives header or footer text without its own font the workbook body font.</summary>
        private static HeaderFooterLineStyle? WithBodyFont(HeaderFooterLineStyle? style, ExcelPdfExportOptions options) {
            if (string.IsNullOrWhiteSpace(options.RunningLineFontFamily) ||
                options.PdfOptions?.HeaderFontFamily != null || options.PdfOptions?.FooterFontFamily != null) return style;
            // Text with its own font, or bold or italic text, keeps the existing font mapping.
            if (style != null && (!string.IsNullOrWhiteSpace(style.FontFamilyName) || style.FontFamily.HasValue || style.Bold || style.Italic)) return style;
            return new HeaderFooterLineStyle {
                FontSize = style?.FontSize ?? options.RunningLineFontSize,
                Color = style?.Color,
                FontFamilyName = options.RunningLineFontFamily
            };
        }

        private static double GetHeaderFooterMarginPoints(ExcelSheetPageSetup? pageSetup, bool header) {
            double inches = header ? pageSetup?.Margins?.Header ?? ExcelNormalHeaderFooterMarginInches : pageSetup?.Margins?.Footer ?? ExcelNormalHeaderFooterMarginInches;
            if (double.IsNaN(inches) || double.IsInfinity(inches) || inches < 0D) inches = ExcelNormalHeaderFooterMarginInches;
            return inches * 72D;
        }

        /// <summary>
        /// Places the running line the way Excel does: the header text starts at the header
        /// margin below the top edge and the footer text ends at the footer margin above the
        /// bottom edge. The PDF line offset is measured from the body edge to the baseline.
        /// A picture that needs more room than that keeps its larger offset.
        /// </summary>
        private static double ResolveHeaderFooterLineOffset(bool header, double bodyMargin, double headerFooterMargin, double? fontSize, double? imageOffset) {
            double size = fontSize is double value && value > 0D ? value : DefaultHeaderFooterFontSize;
            // Typical ascent and descent of office text faces, as a share of the font size.
            double baselineFromEdge = headerFooterMargin + size * (header ? 0.8D : 0.2D);
            double offset = bodyMargin - baselineFromEdge;
            if (imageOffset.HasValue) offset = Math.Max(offset, imageOffset.Value);
            double textAllowance = header ? size * 0.8D : size * 0.2D;
            double maximum = Math.Max(0D, bodyMargin - textAllowance);
            if (offset > maximum) offset = maximum;
            return Math.Max(0D, offset);
        }

        /// <summary>Room left above a header baseline for the text's ascent, in points.</summary>
        private const double HeaderTextAllowance = 10D;

        /// <summary>
        /// The header line starts at the header margin and is as tall as its tallest picture or
        /// its text ascent; the pictures stand on its baseline beside the text.
        /// </summary>
        private static double ResolveHeaderLineOffset(double bodyMargin, double headerMargin, double? fontSize, double tallestImage) {
            double size = fontSize is double value && value > 0D ? value : DefaultHeaderFooterFontSize;
            double baselineFromEdge = headerMargin + Math.Max(size * 0.8D, tallestImage);
            double offset = bodyMargin - baselineFromEdge;
            double maximum = Math.Max(0D, bodyMargin - size * 0.8D);
            return Math.Max(0D, Math.Min(maximum, offset));
        }

        /// <summary>Font size of a header or footer line without explicit formatting, in points.</summary>
        private const double DefaultHeaderFooterFontSize = 9D;

        /// <summary>Default distance of the running header and footer lines from the body, in points.</summary>
        private const double DefaultHeaderFooterOffset = 18D;

        /// <summary>
        /// Excel keeps header and footer pictures in the page margin. The running line hangs
        /// a picture from its baseline toward the body, so move the line out far enough for
        /// the picture to end before the sheet, and shrink a picture taller than the margin.
        /// </summary>
        /// <returns>The line offset to use, or null to keep the default.</returns>
        private static double? FitHeaderFooterImagesToMargin(PreparedHeaderFooterImages images, bool header, double margin, string sheetName, ExcelPdfExportOptions options) {
            const double gap = 2D;
            // Leave room for the text that shares the line: its ascent above a header
            // baseline, its descent below a footer baseline.
            double textAllowance = header ? 10D : 4D;
            double maximumHeight = Math.Max(4D, margin - textAllowance - gap);
            double tallest = 0D;
            bool scaled = false;
            PreparedHeaderFooterImage? Fit(PreparedHeaderFooterImage? image) {
                if (image == null) return null;
                if (image.HeightPoints > maximumHeight) {
                    double factor = maximumHeight / image.HeightPoints;
                    image = new PreparedHeaderFooterImage(image.Bytes, image.WidthPoints * factor, maximumHeight);
                    scaled = true;
                }

                tallest = Math.Max(tallest, image.HeightPoints);
                return image;
            }

            if (header) {
                images.HeaderLeft = Fit(images.HeaderLeft);
                images.HeaderCenter = Fit(images.HeaderCenter);
                images.HeaderRight = Fit(images.HeaderRight);
            } else {
                images.FooterLeft = Fit(images.FooterLeft);
                images.FooterCenter = Fit(images.FooterCenter);
                images.FooterRight = Fit(images.FooterRight);
            }

            if (scaled) {
                AddWarning(options, sheetName, header ? "header-image" : "footer-image",
                    $"A {(header ? "header" : "footer")} picture is taller than the page margin and was scaled down to fit it instead of covering the sheet.");
            }

            if (tallest + gap <= DefaultHeaderFooterOffset) return null;
            return Math.Min(tallest + gap, Math.Max(0D, margin - textAllowance));
        }

        private sealed class PreparedHeaderFooterImages {
            internal PreparedHeaderFooterImage? HeaderLeft { get; set; }
            internal PreparedHeaderFooterImage? HeaderCenter { get; set; }
            internal PreparedHeaderFooterImage? HeaderRight { get; set; }
            internal PreparedHeaderFooterImage? FooterLeft { get; set; }
            internal PreparedHeaderFooterImage? FooterCenter { get; set; }
            internal PreparedHeaderFooterImage? FooterRight { get; set; }

            internal bool HasHeaderImages => HeaderLeft != null || HeaderCenter != null || HeaderRight != null;
            internal bool HasFooterImages => FooterLeft != null || FooterCenter != null || FooterRight != null;
        }

        private sealed class PreparedHeaderFooterImage {
            internal PreparedHeaderFooterImage(byte[] bytes, double widthPoints, double heightPoints) {
                Bytes = bytes;
                WidthPoints = widthPoints;
                HeightPoints = heightPoints;
            }

            internal byte[] Bytes { get; }
            internal double WidthPoints { get; }
            internal double HeightPoints { get; }
        }

    }
}
