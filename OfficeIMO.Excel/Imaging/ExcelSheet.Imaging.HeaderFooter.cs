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
            ExcelWorksheetImageExportOptions options) {
            if (!TryCreateHeaderFooterTextChrome(out HeaderFooterTextChrome chrome)) {
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

        private bool CanRenderHeaderFooterTextChrome() {
            if (!HasHeaderFooterContent()) {
                return true;
            }

            return TryCreateHeaderFooterTextChrome(out _);
        }

        private bool TryCreateHeaderFooterTextChrome(out HeaderFooterTextChrome chrome) {
            chrome = default;
            HeaderFooterSnapshot snapshot = GetHeaderFooter();
            if (snapshot.DifferentFirstPage ||
                snapshot.DifferentOddEven ||
                snapshot.HeaderHasPicturePlaceholder ||
                snapshot.FooterHasPicturePlaceholder ||
                snapshot.HeaderLeftImage != null ||
                snapshot.HeaderCenterImage != null ||
                snapshot.HeaderRightImage != null ||
                snapshot.FooterLeftImage != null ||
                snapshot.FooterCenterImage != null ||
                snapshot.FooterRightImage != null) {
                return false;
            }

            if (!TryNormalizePlainHeaderFooterText(snapshot.HeaderLeft, out string headerLeft) ||
                !TryNormalizePlainHeaderFooterText(snapshot.HeaderCenter, out string headerCenter) ||
                !TryNormalizePlainHeaderFooterText(snapshot.HeaderRight, out string headerRight) ||
                !TryNormalizePlainHeaderFooterText(snapshot.FooterLeft, out string footerLeft) ||
                !TryNormalizePlainHeaderFooterText(snapshot.FooterCenter, out string footerCenter) ||
                !TryNormalizePlainHeaderFooterText(snapshot.FooterRight, out string footerRight)) {
                return false;
            }

            chrome = new HeaderFooterTextChrome(
                headerLeft,
                headerCenter,
                headerRight,
                footerLeft,
                footerCenter,
                footerRight);
            return chrome.HasAnyText;
        }

        private static bool TryNormalizePlainHeaderFooterText(string? text, out string normalized) {
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

                return false;
            }

            normalized = builder.ToString().Trim();
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
    }
}
