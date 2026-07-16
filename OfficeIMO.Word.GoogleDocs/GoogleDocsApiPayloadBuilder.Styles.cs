using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.GoogleDocs {
    internal static partial class GoogleDocsApiPayloadBuilder {
        private static string ResolveListPreset(
            GoogleDocsParagraph paragraph,
            TranslationReport report) {
            if (paragraph.IsOrderedList == true) {
                return "NUMBERED_DECIMAL_NESTED";
            }

            if (paragraph.IsOrderedList == false) {
                return "BULLET_DISC_CIRCLE_SQUARE";
            }

            AddReportNoticeOnce(
                report,
                TranslationSeverity.Warning,
                "Lists",
                "A Word list item did not expose an ordered-vs-bulleted classification, so Google Docs export currently falls back to a bullet preset for that paragraph.");

            return "BULLET_DISC_CIRCLE_SQUARE";
        }

        private static bool TryBuildParagraphDimension(double? points, out GoogleDocsApiDimensionPayload? dimension) {
            if (!points.HasValue) {
                dimension = null;
                return false;
            }

            dimension = new GoogleDocsApiDimensionPayload {
                Magnitude = points.Value,
                Unit = "PT",
            };
            return true;
        }

        private static string ResolveSectionBreakType(
            string? sectionBreakType,
            TranslationReport report) {
            if (string.IsNullOrWhiteSpace(sectionBreakType)) {
                return "NEXT_PAGE";
            }

            switch (sectionBreakType!.Trim().ToUpperInvariant()) {
                case "CONTINUOUS":
                    return "CONTINUOUS";
                case "NEXTPAGE":
                    return "NEXT_PAGE";
                case "EVENPAGE":
                case "ODDPAGE":
                    AddReportNoticeOnce(
                        report,
                        TranslationSeverity.Warning,
                        "SectionBreaks",
                        "Word even-page and odd-page section breaks currently fall back to Google Docs NEXT_PAGE section breaks.");
                    return "NEXT_PAGE";
                default:
                    AddReportNoticeOnce(
                        report,
                        TranslationSeverity.Warning,
                        "SectionBreaks",
                        $"Word section break type '{sectionBreakType}' is not mapped directly yet, so Google Docs export currently falls back to NEXT_PAGE.");
                    return "NEXT_PAGE";
            }
        }

        private static GoogleDocsApiSizePayload? BuildImageSize(GoogleDocsInlineImage image) {
            var width = TryConvertImageDimension(image.Width);
            var height = TryConvertImageDimension(image.Height);
            if (width == null && height == null) {
                return null;
            }

            return new GoogleDocsApiSizePayload {
                Width = width,
                Height = height,
            };
        }

        private static GoogleDocsApiDimensionPayload? TryConvertImageDimension(double? value) {
            if (!value.HasValue || value.Value <= 0) {
                return null;
            }

            // OfficeIMO image dimensions are authored in inches, so translate to Google Docs points.
            return new GoogleDocsApiDimensionPayload {
                Magnitude = Math.Round(value.Value * 72d, 2),
                Unit = "PT",
            };
        }


        private static bool TryMapNamedStyle(GoogleDocsParagraph paragraph, out string namedStyleType) {
            namedStyleType = string.Empty;
            var value = paragraph.StyleId ?? paragraph.StyleName;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            var normalizedValue = value!.Replace(" ", string.Empty).ToUpperInvariant();
            switch (normalizedValue) {
                case "TITLE":
                    namedStyleType = "TITLE";
                    return true;
                case "SUBTITLE":
                    namedStyleType = "SUBTITLE";
                    return true;
                case "NORMAL":
                case "NORMALTEXT":
                    namedStyleType = "NORMAL_TEXT";
                    return true;
                case "HEADING1":
                    namedStyleType = "HEADING_1";
                    return true;
                case "HEADING2":
                    namedStyleType = "HEADING_2";
                    return true;
                case "HEADING3":
                    namedStyleType = "HEADING_3";
                    return true;
                case "HEADING4":
                    namedStyleType = "HEADING_4";
                    return true;
                case "HEADING5":
                    namedStyleType = "HEADING_5";
                    return true;
                case "HEADING6":
                    namedStyleType = "HEADING_6";
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryMapAlignment(string? alignment, out string docsAlignment) {
            docsAlignment = string.Empty;
            if (string.IsNullOrWhiteSpace(alignment)) {
                return false;
            }

            var normalizedAlignment = alignment!.Trim().ToUpperInvariant();
            switch (normalizedAlignment) {
                case "CENTER":
                    docsAlignment = "CENTER";
                    return true;
                case "BOTH":
                case "JUSTIFIED":
                    docsAlignment = "JUSTIFIED";
                    return true;
                case "RIGHT":
                case "END":
                    docsAlignment = "END";
                    return true;
                case "LEFT":
                case "START":
                    docsAlignment = "START";
                    return true;
                default:
                    return false;
            }
        }

        private static GoogleDocsApiOptionalColorPayload? BuildOptionalColor(string? colorHex) {
            if (!TryParseRgbColor(colorHex, out var red, out var green, out var blue)) {
                return null;
            }

            return new GoogleDocsApiOptionalColorPayload {
                Color = new GoogleDocsApiColorPayload {
                    RgbColor = new GoogleDocsApiRgbColorPayload {
                        Red = red,
                        Green = green,
                        Blue = blue,
                    }
                }
            };
        }

        private static GoogleDocsApiOptionalColorPayload? BuildHighlightColor(string? highlightColor) {
            if (string.IsNullOrWhiteSpace(highlightColor)) {
                return null;
            }

            string nonNullHighlightColor = highlightColor!;
            var normalizedHighlightColor = nonNullHighlightColor.Trim().ToUpperInvariant();
            string? colorHex = normalizedHighlightColor switch {
                "YELLOW" => "FFFF00",
                "GREEN" => "00FF00",
                "CYAN" => "00FFFF",
                "MAGENTA" => "FF00FF",
                "BLUE" => "0000FF",
                "RED" => "FF0000",
                "DARKBLUE" => "000080",
                "DARKCYAN" => "008080",
                "DARKGREEN" => "008000",
                "DARKMAGENTA" => "800080",
                "DARKRED" => "800000",
                "DARKYELLOW" => "808000",
                "DARKGRAY" => "808080",
                "LIGHTGRAY" => "D3D3D3",
                "BLACK" => "000000",
                "WHITE" => "FFFFFF",
                "NONE" => null,
                _ => null,
            };

            return BuildOptionalColor(colorHex);
        }

        private static string? BuildBaselineOffset(string? verticalTextAlignment) {
            if (string.IsNullOrWhiteSpace(verticalTextAlignment)) {
                return null;
            }

            string nonNullVerticalTextAlignment = verticalTextAlignment!;
            return nonNullVerticalTextAlignment.Trim().ToUpperInvariant() switch {
                "SUPERSCRIPT" => "SUPERSCRIPT",
                "SUBSCRIPT" => "SUBSCRIPT",
                _ => null,
            };
        }

        private static GoogleDocsApiTableCellBorderPayload? BuildTableCellBorder(GoogleDocsTableCellBorder? border) {
            if (border == null) {
                return null;
            }

            bool isExplicitlyNone = string.Equals(border.Style, "Nil", StringComparison.OrdinalIgnoreCase)
                || string.Equals(border.Style, "None", StringComparison.OrdinalIgnoreCase);

            var width = BuildBorderWidth(border.Size, isExplicitlyNone);
            var color = BuildOptionalColor(border.ColorHex);
            var dashStyle = ResolveTableBorderDashStyle(border.Style);
            if (width == null && color == null && dashStyle == null) {
                return null;
            }

            return new GoogleDocsApiTableCellBorderPayload {
                Width = width,
                Color = color,
                DashStyle = dashStyle,
            };
        }

        private static GoogleDocsApiParagraphBorderPayload? BuildParagraphBorder(GoogleDocsParagraphBorder? border) {
            if (border == null) {
                return null;
            }

            bool isExplicitlyNone = string.Equals(border.Style, "Nil", StringComparison.OrdinalIgnoreCase)
                || string.Equals(border.Style, "None", StringComparison.OrdinalIgnoreCase);

            var width = BuildBorderWidth(border.Size, isExplicitlyNone);
            var color = BuildOptionalColor(border.ColorHex);
            var padding = BuildParagraphBorderPadding(border.Space);
            var dashStyle = ResolveTableBorderDashStyle(border.Style);
            if (width == null && color == null && padding == null && dashStyle == null) {
                return null;
            }

            return new GoogleDocsApiParagraphBorderPayload {
                Width = width,
                Color = color,
                Padding = padding,
                DashStyle = dashStyle,
            };
        }

        private static void AppendSectionStyleRequest(
            GoogleDocsApiBatchUpdatePayload payload,
            int startIndex,
            int endIndex,
            string? segmentId,
            GoogleDocsSectionStyle? sectionStyle) {
            if (payload == null || sectionStyle == null || !string.IsNullOrWhiteSpace(segmentId)) {
                return;
            }

            var stylePayload = BuildSectionStyle(sectionStyle);
            if (stylePayload == null) {
                return;
            }

            var fields = BuildSectionStyleFields(stylePayload);
            if (fields.Count == 0) {
                return;
            }

            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                UpdateSectionStyle = new GoogleDocsApiUpdateSectionStyleRequestPayload {
                    Range = new GoogleDocsApiRangePayload {
                        StartIndex = startIndex,
                        EndIndex = Math.Max(startIndex + 1, endIndex),
                    },
                    SectionStyle = stylePayload,
                    Fields = string.Join(",", fields),
                }
            });
        }

        private static GoogleDocsApiSectionStylePayload? BuildSectionStyle(GoogleDocsSectionStyle source) {
            if (source == null) {
                return null;
            }

            var payload = new GoogleDocsApiSectionStylePayload();

            if (TryBuildParagraphDimension(source.MarginTopPoints, out var marginTop)) {
                payload.MarginTop = marginTop;
            }

            if (TryBuildParagraphDimension(source.MarginBottomPoints, out var marginBottom)) {
                payload.MarginBottom = marginBottom;
            }

            if (TryBuildParagraphDimension(source.MarginLeftPoints, out var marginLeft)) {
                payload.MarginLeft = marginLeft;
            }

            if (TryBuildParagraphDimension(source.MarginRightPoints, out var marginRight)) {
                payload.MarginRight = marginRight;
            }

            if (TryBuildParagraphDimension(source.HeaderMarginPoints, out var marginHeader)) {
                payload.MarginHeader = marginHeader;
            }

            if (TryBuildParagraphDimension(source.FooterMarginPoints, out var marginFooter)) {
                payload.MarginFooter = marginFooter;
            }

            if (BuildSectionColumnProperties(source.ColumnCount, source.ColumnSpacingPoints) is { Count: > 0 } columnProperties) {
                payload.ColumnProperties = columnProperties;
            }

            if (source.ColumnCount.GetValueOrDefault() > 1 || source.HasColumnSeparator) {
                payload.ColumnSeparatorStyle = source.HasColumnSeparator ? "BETWEEN_EACH_COLUMN" : "NONE";
            }

            if (source.UseFirstPageHeaderFooter) {
                payload.UseFirstPageHeaderFooter = true;
            }

            if (source.PageNumberStart.HasValue && source.PageNumberStart.Value > 0) {
                payload.PageNumberStart = source.PageNumberStart.Value;
            }

            if (string.Equals(source.Orientation, "Landscape", StringComparison.OrdinalIgnoreCase)) {
                payload.FlipPageOrientation = true;
            }

            return payload;
        }

        private static List<string> BuildSectionStyleFields(GoogleDocsApiSectionStylePayload style) {
            var fields = new List<string>();
            if (style.MarginTop != null) fields.Add("marginTop");
            if (style.MarginBottom != null) fields.Add("marginBottom");
            if (style.MarginLeft != null) fields.Add("marginLeft");
            if (style.MarginRight != null) fields.Add("marginRight");
            if (style.MarginHeader != null) fields.Add("marginHeader");
            if (style.MarginFooter != null) fields.Add("marginFooter");
            if (style.ColumnProperties != null) fields.Add("columnProperties");
            if (!string.IsNullOrWhiteSpace(style.ColumnSeparatorStyle)) fields.Add("columnSeparatorStyle");
            if (style.UseFirstPageHeaderFooter.HasValue) fields.Add("useFirstPageHeaderFooter");
            if (style.PageNumberStart.HasValue) fields.Add("pageNumberStart");
            if (style.FlipPageOrientation.HasValue) fields.Add("flipPageOrientation");
            return fields;
        }

        private static bool TryBuildSize(double? widthPoints, double? heightPoints, out GoogleDocsApiSizePayload size) {
            size = null!;
            if (!widthPoints.HasValue || !heightPoints.HasValue || widthPoints.Value <= 0 || heightPoints.Value <= 0) {
                return false;
            }

            size = new GoogleDocsApiSizePayload {
                Width = new GoogleDocsApiDimensionPayload {
                    Magnitude = Math.Round(widthPoints.Value, 2, MidpointRounding.AwayFromZero),
                    Unit = "PT",
                },
                Height = new GoogleDocsApiDimensionPayload {
                    Magnitude = Math.Round(heightPoints.Value, 2, MidpointRounding.AwayFromZero),
                    Unit = "PT",
                },
            };
            return true;
        }

        private static void AppendDocumentStyleRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            GoogleDocsBatch batch) {
            if (payload == null || batch == null) {
                return;
            }

            var documentStyle = new GoogleDocsApiDocumentStylePayload();
            var fields = new List<string>();
            var firstSection = batch.Snapshot.Sections.FirstOrDefault();
            if (firstSection != null
                && TryBuildDocumentPageSize(
                    firstSection.PageWidthPoints,
                    firstSection.PageHeightPoints,
                    firstSection.Orientation,
                    out var pageSize)) {
                documentStyle.PageSize = pageSize;
                fields.Add("pageSize");

                bool hasDifferentPaperSize = batch.Snapshot.Sections.Skip(1).Any(section =>
                    !HasSameDocumentPageSize(
                        pageSize,
                        section.PageWidthPoints,
                        section.PageHeightPoints,
                        section.Orientation));
                if (hasDifferentPaperSize) {
                    AddReportNoticeOnce(
                        batch.Report,
                        TranslationSeverity.Warning,
                        "SectionLayout",
                        "Google Docs uses one document-wide page size. The first Word section's paper size is preserved, while later sections with a different paper size retain their orientation and margins but use that document-wide size.");
                }
            }

            bool useEvenPageHeaderFooter = batch.Snapshot.Sections.Any(section => section.DifferentOddAndEvenPages)
                || batch.Segments.Any(segment => string.Equals(segment.Variant, "even", StringComparison.OrdinalIgnoreCase));
            if (useEvenPageHeaderFooter) {
                documentStyle.UseEvenPageHeaderFooter = true;
                fields.Add("useEvenPageHeaderFooter");
            }

            if (fields.Count == 0) {
                return;
            }

            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                UpdateDocumentStyle = new GoogleDocsApiUpdateDocumentStyleRequestPayload {
                    DocumentStyle = documentStyle,
                    Fields = string.Join(",", fields),
                }
            });
        }

        private static bool TryBuildDocumentPageSize(
            double? widthPoints,
            double? heightPoints,
            string? orientation,
            out GoogleDocsApiSizePayload size) {
            if (string.Equals(orientation, "Landscape", StringComparison.OrdinalIgnoreCase)) {
                return TryBuildSize(heightPoints, widthPoints, out size);
            }

            return TryBuildSize(widthPoints, heightPoints, out size);
        }

        private static bool HasSameDocumentPageSize(
            GoogleDocsApiSizePayload expected,
            double? widthPoints,
            double? heightPoints,
            string? orientation) {
            if (!TryBuildDocumentPageSize(widthPoints, heightPoints, orientation, out var actual)) {
                return true;
            }

            return Math.Abs(expected.Width!.Magnitude - actual.Width!.Magnitude) < 0.01d
                && Math.Abs(expected.Height!.Magnitude - actual.Height!.Magnitude) < 0.01d;
        }

        private static List<GoogleDocsApiSectionColumnPropertiesPayload>? BuildSectionColumnProperties(int? columnCount, double? columnSpacingPoints) {
            int count = columnCount.GetValueOrDefault();
            if (count <= 1) {
                return null;
            }

            var columns = new List<GoogleDocsApiSectionColumnPropertiesPayload>();
            for (int index = 0; index < count; index++) {
                var column = new GoogleDocsApiSectionColumnPropertiesPayload();
                if (index < count - 1 && TryBuildParagraphDimension(columnSpacingPoints, out var paddingEnd)) {
                    column.PaddingEnd = paddingEnd;
                }

                columns.Add(column);
            }

            return columns;
        }

        private static List<GoogleDocsApiTabStopPayload>? BuildParagraphTabStops(IReadOnlyList<GoogleDocsTabStop> tabStops) {
            if (tabStops == null || tabStops.Count == 0) {
                return null;
            }

            var result = new List<GoogleDocsApiTabStopPayload>();
            foreach (var tabStop in tabStops) {
                if (!TryBuildParagraphTabStop(tabStop, out var payload)) {
                    continue;
                }

                result.Add(payload);
            }

            return result.Count == 0 ? null : result;
        }

        private static bool TryBuildParagraphTabStop(GoogleDocsTabStop tabStop, out GoogleDocsApiTabStopPayload payload) {
            payload = null!;
            if (tabStop == null || tabStop.OffsetPoints < 0) {
                return false;
            }

            payload = new GoogleDocsApiTabStopPayload {
                Alignment = ResolveParagraphTabStopAlignment(tabStop.Alignment),
                Offset = new GoogleDocsApiDimensionPayload {
                    Magnitude = Math.Round(tabStop.OffsetPoints, 2, MidpointRounding.AwayFromZero),
                    Unit = "PT",
                }
            };

            return true;
        }

        private static string? ResolveParagraphTabStopAlignment(string? alignment) {
            if (string.IsNullOrWhiteSpace(alignment)) {
                return null;
            }

            switch (alignment!.Trim().ToUpperInvariant()) {
                case "LEFT":
                case "START":
                case "BAR":
                case "CLEAR":
                case "LIST":
                    return "START";
                case "CENTER":
                    return "CENTER";
                case "RIGHT":
                case "END":
                    return "END";
                case "DECIMAL":
                    return "DECIMAL";
                default:
                    return "START";
            }
        }

        private static GoogleDocsApiDimensionPayload? BuildBorderWidth(uint? size, bool isExplicitlyNone) {
            if (isExplicitlyNone) {
                return new GoogleDocsApiDimensionPayload {
                    Magnitude = 0,
                    Unit = "PT",
                };
            }

            if (!size.HasValue || size.Value == 0) {
                return null;
            }

            return new GoogleDocsApiDimensionPayload {
                Magnitude = Math.Round(size.Value / 8d, 2, MidpointRounding.AwayFromZero),
                Unit = "PT",
            };
        }

        private static GoogleDocsApiDimensionPayload? BuildParagraphBorderPadding(uint? space) {
            if (!space.HasValue) {
                return null;
            }

            return new GoogleDocsApiDimensionPayload {
                Magnitude = space.Value,
                Unit = "PT",
            };
        }

        private static string? ResolveTableBorderDashStyle(string? style) {
            if (string.IsNullOrWhiteSpace(style)) {
                return null;
            }

            switch (style!.Trim().ToUpperInvariant()) {
                case "NONE":
                case "NIL":
                case "SINGLE":
                case "THICK":
                case "DOUBLE":
                case "TRIPLE":
                case "THINTHICKSMALLGAP":
                case "THICKTHINSMALLGAP":
                case "THINTHICKTHINSMALLGAP":
                case "THINTHICKMEDIUMGAP":
                case "THICKTHINMEDIUMGAP":
                case "THINTHICKTHINMEDIUMGAP":
                case "THINTHICKLARGEGAP":
                case "THICKTHINLARGEGAP":
                case "THINTHICKTHINLARGEGAP":
                case "WAVE":
                case "DOUBLEWAVE":
                case "THREED":
                case "THREEDEMBOSS":
                case "THREEDENGRAVE":
                case "OUTSET":
                case "INSET":
                    return "SOLID";
                case "DASHDOT":
                case "DASHDOTSTROKED":
                case "DOTDASH":
                case "DOTDOTDASH":
                    return "DASH_DOT";
                case "DASH":
                case "DASHED":
                case "DASHSMALLGAP":
                case "DASHDOTDOTHEAVY":
                case "DASHDOTHEAVY":
                case "DASHLONG":
                case "DASHLONGHEAVY":
                    return "DASH";
                case "DOT":
                case "DOTTED":
                case "DOTTEDDASH":
                case "DASHEDHEAVY":
                case "DOTTEDHEAVY":
                    return "DOT";
                default:
                    return "SOLID";
            }
        }

        private static bool TryParseRgbColor(string? colorHex, out double red, out double green, out double blue) {
            red = green = blue = 0;
            if (string.IsNullOrWhiteSpace(colorHex)) {
                return false;
            }

            var normalized = colorHex!.Trim().TrimStart('#');
            if (normalized.Length == 8) {
                normalized = normalized.Substring(2);
            }

            if (normalized.Length != 6) {
                return false;
            }

            if (!int.TryParse(normalized.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var redByte)
                || !int.TryParse(normalized.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var greenByte)
                || !int.TryParse(normalized.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var blueByte)) {
                return false;
            }

            red = redByte / 255d;
            green = greenByte / 255d;
            blue = blueByte / 255d;
            return true;
        }

        private static string SanitizeText(string? value) {
            if (string.IsNullOrEmpty(value)) {
                return string.Empty;
            }

            var normalized = value!.Replace("\r\n", "\n").Replace('\r', '\n');
            var builder = new StringBuilder(normalized.Length);
            foreach (var character in normalized) {
                if (character == '\n' || character == '\t' || !char.IsControl(character)) {
                    builder.Append(character);
                }
            }

            return builder.ToString();
        }

        private static void AddReportNoticeOnce(
            TranslationReport report,
            TranslationSeverity severity,
            string feature,
            string message) {
            if (report.Notices.Any(notice =>
                notice.Severity == severity
                && string.Equals(notice.Feature, feature, StringComparison.Ordinal)
                && string.Equals(notice.Message, message, StringComparison.Ordinal))) {
                return;
            }

            report.Add(severity, feature, message);
        }
    }
}
