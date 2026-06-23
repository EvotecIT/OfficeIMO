using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const double HeaderFooterBandHeight = 28D;
        private const double HeaderFooterFontSize = 12D;
        private const double HeaderFooterHorizontalPadding = 8D;
        private const double HeaderFooterZoneGap = 4D;
        private const string HeaderFooterFontFamily = "Arial, sans-serif";
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

            string headerFooterSource = Name + "!headerFooter";
            var combinedDiagnostics = new List<OfficeImageExportDiagnostic>(content.Diagnostics.Count + 2);
            combinedDiagnostics.AddRange(content.Diagnostics);
            if (chrome.HasFormatting) {
                combinedDiagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Info,
                    ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation,
                    "Worksheet header/footer text formatting was rendered through the dependency-free image approximation path.",
                    headerFooterSource));
            }

            AddHeaderFooterFontFamilyFallbackDiagnostics(chrome, headerFooterSource, combinedDiagnostics);
            IReadOnlyList<OfficeImageExportDiagnostic> diagnostics = combinedDiagnostics.Count == content.Diagnostics.Count
                ? content.Diagnostics
                : combinedDiagnostics.AsReadOnly();

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
                    diagnostics);
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
                diagnostics);
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

        private static void AddHeaderFooterFontFamilyFallbackDiagnostics(HeaderFooterTextChrome chrome, string source, List<OfficeImageExportDiagnostic> diagnostics) {
            var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddHeaderFooterFontFamilyFallbackDiagnostics(chrome.HeaderLeft, source, diagnostics, reported);
            AddHeaderFooterFontFamilyFallbackDiagnostics(chrome.HeaderCenter, source, diagnostics, reported);
            AddHeaderFooterFontFamilyFallbackDiagnostics(chrome.HeaderRight, source, diagnostics, reported);
            AddHeaderFooterFontFamilyFallbackDiagnostics(chrome.FooterLeft, source, diagnostics, reported);
            AddHeaderFooterFontFamilyFallbackDiagnostics(chrome.FooterCenter, source, diagnostics, reported);
            AddHeaderFooterFontFamilyFallbackDiagnostics(chrome.FooterRight, source, diagnostics, reported);
        }

        private static void AddHeaderFooterFontFamilyFallbackDiagnostics(
            HeaderFooterTextSection section,
            string source,
            List<OfficeImageExportDiagnostic> diagnostics,
            HashSet<string> reported) {
            for (int index = 0; index < section.Runs.Count; index++) {
                string? fontFamily = section.Runs[index].FontFamily;
                if (string.IsNullOrWhiteSpace(fontFamily) || !reported.Add(fontFamily!) || OfficeTrueTypeFont.TryLoadFontFamily(fontFamily, out _) != null) {
                    continue;
                }

                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.HeaderFooterFontFamilyFallback,
                    "Worksheet header/footer font family '" + fontFamily + "' could not be loaded exactly by the dependency-free image exporter; raster text metrics and image output used the shared fallback font path.",
                    source));
            }
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
            if (!TryResolveHeaderFooterText(headerLeftSource, pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextSection headerLeft) ||
                !TryResolveHeaderFooterText(headerCenterSource, pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextSection headerCenter) ||
                !TryResolveHeaderFooterText(headerRightSource, pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextSection headerRight) ||
                !TryResolveHeaderFooterText(footerLeftSource, pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextSection footerLeft) ||
                !TryResolveHeaderFooterText(footerCenterSource, pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextSection footerCenter) ||
                !TryResolveHeaderFooterText(footerRightSource, pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextSection footerRight)) {
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
            OfficeTextZoneLayout zones = OfficeTextZoneLayout.CreateThreeColumn(width, padding, HeaderFooterZoneGap * scale);
            if (chrome.HasHeader) {
                double y = Math.Max(0D, (headerHeight - fontSize) / 2D);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderLeft, zones.Left, y, fontSize, OfficeTextAlignment.Left);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderCenter, zones.Center, y, fontSize, OfficeTextAlignment.Center);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderRight, zones.Right, y, fontSize, OfficeTextAlignment.Right);
            }

            if (chrome.HasFooter) {
                double y = height - footerHeight + Math.Max(0D, (footerHeight - fontSize) / 2D);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterLeft, zones.Left, y, fontSize, OfficeTextAlignment.Left);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterCenter, zones.Center, y, fontSize, OfficeTextAlignment.Center);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterRight, zones.Right, y, fontSize, OfficeTextAlignment.Right);
            }
        }

        private static void DrawHeaderFooterRasterLine(
            OfficeRasterCanvas canvas,
            HeaderFooterTextSection section,
            OfficeTextZone zone,
            double y,
            double fontSize,
            OfficeTextAlignment alignment) {
            if (!section.HasText) {
                return;
            }

            using (canvas.PushClipRectangle(zone.X, 0D, zone.Width, canvas.Height)) {
                if (section.HasFormatting) {
                    DrawHeaderFooterRasterRichLine(canvas, section, zone, y, fontSize, alignment);
                } else {
                    string displayText = ResolveHeaderFooterZoneText(section.Text, fontSize, zone.Width, (text, size) => canvas.MeasureText(text, size, HeaderFooterFontFamily), alignment);
                    if (!string.IsNullOrWhiteSpace(displayText)) {
                        canvas.DrawTextLine(displayText, zone.AnchorX, y, fontSize, HeaderFooterTextColor, alignment: alignment, fontFamily: HeaderFooterFontFamily);
                    }
                }
            }
        }

        private static void DrawHeaderFooterRasterRichLine(
            OfficeRasterCanvas canvas,
            HeaderFooterTextSection section,
            OfficeTextZone zone,
            double y,
            double fontSize,
            OfficeTextAlignment alignment) {
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                section.ToOfficeRuns(fontSize, HeaderFooterTextColor),
                zone.Width,
                Math.Ceiling(section.GetMaxResolvedFontSize(fontSize) * 1.2D),
                1.2D,
                (text, size, family) => canvas.MeasureText(text, size, family),
                wrap: false);
            OfficeTextBlockRenderer.DrawRasterRichTextBlock(
                canvas,
                layout,
                zone.X,
                y,
                zone.Width,
                layout.Height,
                alignment);
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
            OfficeTextZoneLayout zones = OfficeTextZoneLayout.CreateThreeColumn(width, padding, HeaderFooterZoneGap * scale);
            var textMeasureCanvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1, OfficeColor.Transparent));
            if (chrome.HasHeader) {
                double baseline = Math.Max(fontSize, (headerHeight + fontSize) / 2D);
                AppendHeaderFooterSvgLine(builder, chrome.HeaderLeft, zones.Left, 0D, headerHeight, baseline, lineHeight, fontSize, OfficeTextAlignment.Left, "header-left", (text, size, family) => textMeasureCanvas.MeasureText(text, size, family));
                AppendHeaderFooterSvgLine(builder, chrome.HeaderCenter, zones.Center, 0D, headerHeight, baseline, lineHeight, fontSize, OfficeTextAlignment.Center, "header-center", (text, size, family) => textMeasureCanvas.MeasureText(text, size, family));
                AppendHeaderFooterSvgLine(builder, chrome.HeaderRight, zones.Right, 0D, headerHeight, baseline, lineHeight, fontSize, OfficeTextAlignment.Right, "header-right", (text, size, family) => textMeasureCanvas.MeasureText(text, size, family));
            }

            if (chrome.HasFooter) {
                double footerTop = height - (chrome.HasFooter ? Math.Max(1, (int)Math.Ceiling(HeaderFooterBandHeight * scale)) : 0);
                double baseline = footerTop + Math.Max(fontSize, ((HeaderFooterBandHeight * scale) + fontSize) / 2D);
                AppendHeaderFooterSvgLine(builder, chrome.FooterLeft, zones.Left, footerTop, height - footerTop, baseline, lineHeight, fontSize, OfficeTextAlignment.Left, "footer-left", (text, size, family) => textMeasureCanvas.MeasureText(text, size, family));
                AppendHeaderFooterSvgLine(builder, chrome.FooterCenter, zones.Center, footerTop, height - footerTop, baseline, lineHeight, fontSize, OfficeTextAlignment.Center, "footer-center", (text, size, family) => textMeasureCanvas.MeasureText(text, size, family));
                AppendHeaderFooterSvgLine(builder, chrome.FooterRight, zones.Right, footerTop, height - footerTop, baseline, lineHeight, fontSize, OfficeTextAlignment.Right, "footer-right", (text, size, family) => textMeasureCanvas.MeasureText(text, size, family));
            }
        }

        private static void AppendHeaderFooterSvgLine(
            StringBuilder builder,
            HeaderFooterTextSection section,
            OfficeTextZone zone,
            double bandTop,
            double bandHeight,
            double baseline,
            double lineHeight,
            double fontSize,
            OfficeTextAlignment alignment,
            string clipSuffix,
            Func<string?, double, string?, double> measure) {
            if (!section.HasText) {
                return;
            }

            string clipId = "xl-header-footer-" + clipSuffix;
            builder.AppendRectClipPathDefinition(clipId, zone.X, bandTop, zone.Width, bandHeight, wrapInDefs: true);
            builder.Append("<g").AppendClipPathReference(clipId).Append(">");
            if (section.HasFormatting) {
                AppendHeaderFooterSvgRichLine(builder, section, zone, baseline, fontSize, lineHeight, alignment, measure);
            } else {
                string displayText = ResolveHeaderFooterZoneText(section.Text, fontSize, zone.Width, (text, size) => measure(text, size, HeaderFooterFontFamily), alignment);
                if (!string.IsNullOrWhiteSpace(displayText)) {
                    builder.AppendSvgTextElement(
                        displayText,
                        zone.AnchorX,
                        baseline,
                        lineHeight,
                        HeaderFooterTextColor,
                        HeaderFooterFontFamily,
                        fontSize,
                        alignment);
                }
            }

            builder.Append("</g>");
        }

        private static void AppendHeaderFooterSvgRichLine(
            StringBuilder builder,
            HeaderFooterTextSection section,
            OfficeTextZone zone,
            double baseline,
            double fontSize,
            double lineHeight,
            OfficeTextAlignment alignment,
            Func<string?, double, string?, double> measure) {
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                section.ToOfficeRuns(fontSize, HeaderFooterTextColor),
                zone.Width,
                Math.Max(lineHeight, section.GetMaxResolvedFontSize(fontSize) * 1.2D),
                1.2D,
                measure,
                wrap: false);
            if (layout.Lines.Count == 0) {
                return;
            }

            OfficeRichTextLine line = layout.Lines[0];
            double cursor = OfficeTextPlacement.ResolveLineLeft(zone.X, zone.Width, line.Width, alignment);
            for (int index = 0; index < line.Segments.Count; index++) {
                OfficeRichTextSegment segment = line.Segments[index];
                builder.AppendSvgRichTextSegment(segment, cursor, baseline);
                cursor += segment.Width;
            }
        }

        private static string ResolveHeaderFooterZoneText(
            string text,
            double fontSize,
            double maxWidth,
            Func<string?, double, double> measure,
            OfficeTextAlignment alignment) {
            if (alignment == OfficeTextAlignment.Right) {
                return OfficeTextLayoutEngine.TrimLineStartToWidth(text, fontSize, maxWidth, measure, out _).Text;
            }

            return OfficeTextLayoutEngine.TrimLineToWidth(text, fontSize, maxWidth, measure, out _).Text;
        }

        private readonly struct HeaderFooterTextChrome {
            internal HeaderFooterTextChrome(
                HeaderFooterTextSection headerLeft,
                HeaderFooterTextSection headerCenter,
                HeaderFooterTextSection headerRight,
                HeaderFooterTextSection footerLeft,
                HeaderFooterTextSection footerCenter,
                HeaderFooterTextSection footerRight) {
                HeaderLeft = headerLeft;
                HeaderCenter = headerCenter;
                HeaderRight = headerRight;
                FooterLeft = footerLeft;
                FooterCenter = footerCenter;
                FooterRight = footerRight;
            }

            internal HeaderFooterTextSection HeaderLeft { get; }
            internal HeaderFooterTextSection HeaderCenter { get; }
            internal HeaderFooterTextSection HeaderRight { get; }
            internal HeaderFooterTextSection FooterLeft { get; }
            internal HeaderFooterTextSection FooterCenter { get; }
            internal HeaderFooterTextSection FooterRight { get; }
            internal bool HasHeader => HeaderLeft.HasText || HeaderCenter.HasText || HeaderRight.HasText;
            internal bool HasFooter => FooterLeft.HasText || FooterCenter.HasText || FooterRight.HasText;
            internal bool HasAnyText => HasHeader || HasFooter;
            internal bool HasFormatting =>
                HeaderLeft.HasFormatting ||
                HeaderCenter.HasFormatting ||
                HeaderRight.HasFormatting ||
                FooterLeft.HasFormatting ||
                FooterCenter.HasFormatting ||
                FooterRight.HasFormatting;
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
