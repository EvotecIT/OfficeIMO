using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const double HeaderFooterBandHeight = 28D;
        private const double HeaderFooterFontSize = 12D;
        private const double HeaderFooterHorizontalPadding = 8D;
        private static readonly OfficeColor HeaderFooterTextColor = OfficeColor.FromRgb(31, 41, 55);

        private OfficeImageExportResult ApplyHeaderFooterTextChrome(
            OfficeImageExportFormat format,
            OfficeImageExportResult content,
            ExcelWorksheetImageExportOptions options,
            int pageNumber,
            int pageCount) {
            DateTime headerFooterDateTime = options.HeaderFooterDateTime ?? DateTime.Now;
            if (!TryCreateHeaderFooterTextChrome(pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextChrome chrome)) {
                return content;
            }

            double scale = options.Scale;
            int headerHeight = chrome.HasHeader ? Math.Max(1, (int)Math.Ceiling(HeaderFooterBandHeight * scale)) : 0;
            int footerHeight = chrome.HasFooter ? Math.Max(1, (int)Math.Ceiling(HeaderFooterBandHeight * scale)) : 0;
            int width = Math.Max(1, content.Width);
            int height = Math.Max(1, content.Height + headerHeight + footerHeight);

            if (format == OfficeImageExportFormat.Svg) {
                OfficeImageLayer layer = OfficeImageLayer.FromSvgInner(
                    OfficeSvgFormatting.ExtractSvgInner(Encoding.UTF8.GetString(content.Bytes)),
                    0D,
                    headerHeight,
                    content.Width,
                    content.Height);
                return new OfficeImageExportResult(
                    format,
                    width,
                    height,
                    OfficeImageComposer.ComposeSvgBytes(
                        width,
                        height,
                        options.BackgroundColor,
                        new[] { layer },
                        beforeLayers: builder => AppendHeaderFooterSvgText(builder, chrome, width, height, headerHeight, options.Scale)),
                    content.Name,
                    content.Source,
                    content.Diagnostics);
            }

            if (!OfficePngReader.TryDecode(content.Bytes, out OfficeRasterImage? contentImage) || contentImage == null) {
                return content;
            }

            OfficeImageLayer contentLayer = OfficeImageLayer.FromRaster(contentImage, 0D, headerHeight, content.Width, content.Height);
            return new OfficeImageExportResult(
                format,
                width,
                height,
                OfficeImageComposer.ComposePng(
                    width,
                    height,
                    options.BackgroundColor,
                    new[] { contentLayer },
                    beforeLayers: canvas => DrawHeaderFooterRaster(canvas, chrome, width, height, headerHeight, footerHeight, scale)),
                content.Name,
                content.Source,
                content.Diagnostics);
        }

        private bool CanRenderHeaderFooterTextChrome(DateTime headerFooterDateTime) {
            if (!HasHeaderFooterContent()) {
                return true;
            }

            HeaderFooterSnapshot snapshot = GetHeaderFooter();
            if (HasUnsupportedHeaderFooterImages(snapshot)) {
                return false;
            }

            if (!TryCreateResolvedHeaderFooterTextChrome(
                snapshot.HeaderLeft,
                snapshot.HeaderCenter,
                snapshot.HeaderRight,
                snapshot.FooterLeft,
                snapshot.FooterCenter,
                snapshot.FooterRight,
                3,
                3,
                headerFooterDateTime,
                out _)) {
                return false;
            }

            if (snapshot.DifferentFirstPage &&
                !TryCreateResolvedHeaderFooterTextChrome(
                    snapshot.FirstHeaderLeft,
                    snapshot.FirstHeaderCenter,
                    snapshot.FirstHeaderRight,
                    snapshot.FirstFooterLeft,
                    snapshot.FirstFooterCenter,
                    snapshot.FirstFooterRight,
                    1,
                    3,
                    headerFooterDateTime,
                    out _)) {
                return false;
            }

            if (snapshot.DifferentOddEven &&
                !TryCreateResolvedHeaderFooterTextChrome(
                    snapshot.EvenHeaderLeft,
                    snapshot.EvenHeaderCenter,
                    snapshot.EvenHeaderRight,
                    snapshot.EvenFooterLeft,
                    snapshot.EvenFooterCenter,
                    snapshot.EvenFooterRight,
                    2,
                    3,
                    headerFooterDateTime,
                    out _)) {
                return false;
            }

            return true;
        }

        private bool TryCreateHeaderFooterTextChrome(int pageNumber, int pageCount, DateTime headerFooterDateTime, out HeaderFooterTextChrome chrome) {
            chrome = default;
            HeaderFooterSnapshot snapshot = GetHeaderFooter();
            if (HasUnsupportedHeaderFooterImages(snapshot)) {
                return false;
            }

            HeaderFooterVariantText selected = SelectHeaderFooterVariantText(snapshot, pageNumber);
            return TryCreateResolvedHeaderFooterTextChrome(
                selected.HeaderLeft,
                selected.HeaderCenter,
                selected.HeaderRight,
                selected.FooterLeft,
                selected.FooterCenter,
                selected.FooterRight,
                pageNumber,
                pageCount,
                headerFooterDateTime,
                out chrome) && chrome.HasAnyText;
        }

        private bool TryCreateResolvedHeaderFooterTextChrome(
            string? headerLeftSource,
            string? headerCenterSource,
            string? headerRightSource,
            string? footerLeftSource,
            string? footerCenterSource,
            string? footerRightSource,
            int pageNumber,
            int pageCount,
            DateTime headerFooterDateTime,
            out HeaderFooterTextChrome chrome) {
            chrome = default;
            if (!TryResolveHeaderFooterText(headerLeftSource, pageNumber, pageCount, headerFooterDateTime, out string headerLeft) ||
                !TryResolveHeaderFooterText(headerCenterSource, pageNumber, pageCount, headerFooterDateTime, out string headerCenter) ||
                !TryResolveHeaderFooterText(headerRightSource, pageNumber, pageCount, headerFooterDateTime, out string headerRight) ||
                !TryResolveHeaderFooterText(footerLeftSource, pageNumber, pageCount, headerFooterDateTime, out string footerLeft) ||
                !TryResolveHeaderFooterText(footerCenterSource, pageNumber, pageCount, headerFooterDateTime, out string footerCenter) ||
                !TryResolveHeaderFooterText(footerRightSource, pageNumber, pageCount, headerFooterDateTime, out string footerRight)) {
                return false;
            }

            chrome = new HeaderFooterTextChrome(
                headerLeft,
                headerCenter,
                headerRight,
                footerLeft,
                footerCenter,
                footerRight);
            return true;
        }

        private static bool HasUnsupportedHeaderFooterImages(HeaderFooterSnapshot snapshot) =>
            snapshot.HeaderHasPicturePlaceholder ||
            snapshot.FooterHasPicturePlaceholder ||
            snapshot.HeaderLeftImage != null ||
            snapshot.HeaderCenterImage != null ||
            snapshot.HeaderRightImage != null ||
            snapshot.FooterLeftImage != null ||
            snapshot.FooterCenterImage != null ||
            snapshot.FooterRightImage != null;

        private static HeaderFooterVariantText SelectHeaderFooterVariantText(HeaderFooterSnapshot snapshot, int pageNumber) {
            if (pageNumber == 1 && snapshot.DifferentFirstPage) {
                return new HeaderFooterVariantText(
                    snapshot.FirstHeaderLeft,
                    snapshot.FirstHeaderCenter,
                    snapshot.FirstHeaderRight,
                    snapshot.FirstFooterLeft,
                    snapshot.FirstFooterCenter,
                    snapshot.FirstFooterRight);
            }

            if (pageNumber % 2 == 0 && snapshot.DifferentOddEven) {
                return new HeaderFooterVariantText(
                    snapshot.EvenHeaderLeft,
                    snapshot.EvenHeaderCenter,
                    snapshot.EvenHeaderRight,
                    snapshot.EvenFooterLeft,
                    snapshot.EvenFooterCenter,
                    snapshot.EvenFooterRight);
            }

            return new HeaderFooterVariantText(
                snapshot.HeaderLeft,
                snapshot.HeaderCenter,
                snapshot.HeaderRight,
                snapshot.FooterLeft,
                snapshot.FooterCenter,
                snapshot.FooterRight);
        }

        private bool TryResolveHeaderFooterText(string? text, int pageNumber, int pageCount, DateTime headerFooterDateTime, out string normalized) {
            normalized = string.Empty;
            if (string.IsNullOrWhiteSpace(text)) {
                return true;
            }

            var builder = new StringBuilder(text!.Length);
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch != '&') {
                    builder.Append(ch);
                    continue;
                }

                if (i + 1 >= text.Length) {
                    return false;
                }

                char token = text[++i];
                if (token == '&') {
                    builder.Append('&');
                } else if (token == 'P') {
                    builder.Append(pageNumber.ToString(CultureInfo.InvariantCulture));
                } else if (token == 'N') {
                    builder.Append(pageCount.ToString(CultureInfo.InvariantCulture));
                } else if (token == 'D') {
                    builder.Append(FormatHeaderFooterDate(headerFooterDateTime));
                } else if (token == 'T') {
                    builder.Append(FormatHeaderFooterTime(headerFooterDateTime));
                } else if (token == 'A') {
                    builder.Append(Name);
                } else if (token == 'F') {
                    if (!TryGetWorkbookFileName(out string fileName)) {
                        return false;
                    }

                    builder.Append(fileName);
                } else if (token == 'Z') {
                    if (!TryGetWorkbookPathPrefix(out string pathPrefix)) {
                        return false;
                    }

                    builder.Append(pathPrefix);
                } else if (token == '[') {
                    int end = text.IndexOf(']', i + 1);
                    if (end < 0) {
                        return false;
                    }

                    string fieldName = text.Substring(i + 1, end - i - 1);
                    if (!TryAppendHeaderFooterField(builder, fieldName, pageNumber, pageCount, headerFooterDateTime)) {
                        return false;
                    }

                    i = end;
                } else {
                    return false;
                }
            }

            normalized = builder.ToString().Trim();
            return true;
        }

        private bool TryAppendHeaderFooterField(StringBuilder builder, string fieldName, int pageNumber, int pageCount, DateTime headerFooterDateTime) {
            if (string.Equals(fieldName, "Page", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(pageNumber.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (string.Equals(fieldName, "Pages", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(pageCount.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (string.Equals(fieldName, "Tab", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(Name);
                return true;
            }

            if (string.Equals(fieldName, "Date", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(FormatHeaderFooterDate(headerFooterDateTime));
                return true;
            }

            if (string.Equals(fieldName, "Time", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(FormatHeaderFooterTime(headerFooterDateTime));
                return true;
            }

            if (string.Equals(fieldName, "File", StringComparison.OrdinalIgnoreCase)) {
                if (!TryGetWorkbookFileName(out string fileName)) {
                    return false;
                }

                builder.Append(fileName);
                return true;
            }

            if (string.Equals(fieldName, "Path", StringComparison.OrdinalIgnoreCase)) {
                if (!TryGetWorkbookPathPrefix(out string pathPrefix)) {
                    return false;
                }

                builder.Append(pathPrefix);
                return true;
            }

            return false;
        }

        private static string FormatHeaderFooterDate(DateTime headerFooterDateTime) =>
            headerFooterDateTime.ToString("d", CultureInfo.CurrentCulture);

        private static string FormatHeaderFooterTime(DateTime headerFooterDateTime) =>
            headerFooterDateTime.ToString("t", CultureInfo.CurrentCulture);

        private bool TryGetWorkbookFileName(out string fileName) {
            fileName = string.Empty;
            string path = _excelDocument.FilePath;
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            fileName = Path.GetFileName(path);
            return !string.IsNullOrWhiteSpace(fileName);
        }

        private bool TryGetWorkbookPathPrefix(out string pathPrefix) {
            pathPrefix = string.Empty;
            string path = _excelDocument.FilePath;
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (string.IsNullOrWhiteSpace(directory)) {
                return true;
            }

            pathPrefix = directory!.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
                directory.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? directory
                : directory + Path.DirectorySeparatorChar;
            return true;
        }

        private static void DrawHeaderFooterRaster(
            OfficeRasterCanvas canvas,
            HeaderFooterTextChrome chrome,
            int width,
            int height,
            int headerHeight,
            int footerHeight,
            double scale) {
            double fontSize = HeaderFooterFontSize * scale;
            double padding = HeaderFooterHorizontalPadding * scale;
            if (chrome.HasHeader) {
                double y = Math.Max(0D, (headerHeight - fontSize) / 2D);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderLeft, padding, y, fontSize, OfficeTextAlignment.Left);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderCenter, width / 2D, y, fontSize, OfficeTextAlignment.Center);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderRight, width - padding, y, fontSize, OfficeTextAlignment.Right);
            }

            if (chrome.HasFooter) {
                double y = height - footerHeight + Math.Max(0D, (footerHeight - fontSize) / 2D);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterLeft, padding, y, fontSize, OfficeTextAlignment.Left);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterCenter, width / 2D, y, fontSize, OfficeTextAlignment.Center);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterRight, width - padding, y, fontSize, OfficeTextAlignment.Right);
            }
        }

        private static void DrawHeaderFooterRasterLine(
            OfficeRasterCanvas canvas,
            string text,
            double x,
            double y,
            double fontSize,
            OfficeTextAlignment alignment) {
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            canvas.DrawTextLine(text, x, y, fontSize, HeaderFooterTextColor, alignment: alignment);
        }

        private static void AppendHeaderFooterSvgText(
            StringBuilder builder,
            HeaderFooterTextChrome chrome,
            int width,
            int height,
            int headerHeight,
            double scale) {
            double fontSize = HeaderFooterFontSize * scale;
            double padding = HeaderFooterHorizontalPadding * scale;
            double lineHeight = fontSize * 1.2D;
            if (chrome.HasHeader) {
                double baseline = Math.Max(fontSize, (headerHeight + fontSize) / 2D);
                AppendHeaderFooterSvgLine(builder, chrome.HeaderLeft, padding, baseline, lineHeight, fontSize, OfficeTextAlignment.Left);
                AppendHeaderFooterSvgLine(builder, chrome.HeaderCenter, width / 2D, baseline, lineHeight, fontSize, OfficeTextAlignment.Center);
                AppendHeaderFooterSvgLine(builder, chrome.HeaderRight, width - padding, baseline, lineHeight, fontSize, OfficeTextAlignment.Right);
            }

            if (chrome.HasFooter) {
                double footerTop = height - (chrome.HasFooter ? Math.Max(1, (int)Math.Ceiling(HeaderFooterBandHeight * scale)) : 0);
                double baseline = footerTop + Math.Max(fontSize, ((HeaderFooterBandHeight * scale) + fontSize) / 2D);
                AppendHeaderFooterSvgLine(builder, chrome.FooterLeft, padding, baseline, lineHeight, fontSize, OfficeTextAlignment.Left);
                AppendHeaderFooterSvgLine(builder, chrome.FooterCenter, width / 2D, baseline, lineHeight, fontSize, OfficeTextAlignment.Center);
                AppendHeaderFooterSvgLine(builder, chrome.FooterRight, width - padding, baseline, lineHeight, fontSize, OfficeTextAlignment.Right);
            }
        }

        private static void AppendHeaderFooterSvgLine(
            StringBuilder builder,
            string text,
            double x,
            double baseline,
            double lineHeight,
            double fontSize,
            OfficeTextAlignment alignment) {
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            builder.AppendSvgTextElement(
                text,
                x,
                baseline,
                lineHeight,
                HeaderFooterTextColor,
                "Arial, sans-serif",
                fontSize,
                alignment);
        }

        private readonly struct HeaderFooterTextChrome {
            internal HeaderFooterTextChrome(
                string headerLeft,
                string headerCenter,
                string headerRight,
                string footerLeft,
                string footerCenter,
                string footerRight) {
                HeaderLeft = headerLeft;
                HeaderCenter = headerCenter;
                HeaderRight = headerRight;
                FooterLeft = footerLeft;
                FooterCenter = footerCenter;
                FooterRight = footerRight;
            }

            internal string HeaderLeft { get; }
            internal string HeaderCenter { get; }
            internal string HeaderRight { get; }
            internal string FooterLeft { get; }
            internal string FooterCenter { get; }
            internal string FooterRight { get; }
            internal bool HasHeader => HasText(HeaderLeft) || HasText(HeaderCenter) || HasText(HeaderRight);
            internal bool HasFooter => HasText(FooterLeft) || HasText(FooterCenter) || HasText(FooterRight);
            internal bool HasAnyText => HasHeader || HasFooter;
            private static bool HasText(string text) => !string.IsNullOrWhiteSpace(text);
        }

        private readonly struct HeaderFooterVariantText {
            internal HeaderFooterVariantText(
                string headerLeft,
                string headerCenter,
                string headerRight,
                string footerLeft,
                string footerCenter,
                string footerRight) {
                HeaderLeft = headerLeft;
                HeaderCenter = headerCenter;
                HeaderRight = headerRight;
                FooterLeft = footerLeft;
                FooterCenter = footerCenter;
                FooterRight = footerRight;
            }

            internal string HeaderLeft { get; }
            internal string HeaderCenter { get; }
            internal string HeaderRight { get; }
            internal string FooterLeft { get; }
            internal string FooterCenter { get; }
            internal string FooterRight { get; }
        }
    }
}
