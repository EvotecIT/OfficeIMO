using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const double HeaderFooterBandHeight = 28D;
        private const double HeaderFooterFontSize = 12D;
        private const double HeaderFooterHorizontalPadding = 8D;
        private const double HeaderFooterZoneGap = 4D;
        private const string HeaderFooterFallbackFontFamily = "Arial, sans-serif";
        private static readonly OfficeColor HeaderFooterTextColor = OfficeColor.FromRgb(31, 41, 55);

        private OfficeImageExportResult ApplyHeaderFooterTextChrome(
            OfficeImageExportFormat format,
            OfficeImageExportFormat rasterPlanningFormat,
            OfficeImageExportResult content,
            ExcelWorksheetImageExportOptions options,
            HeaderFooterSnapshot? headerFooterSnapshot,
            int pageNumber,
            int pageCount,
            ref ExcelRasterRenderState rasterState) {
            DateTime headerFooterDateTime = options.HeaderFooterDateTime ?? DateTime.Now;
            if (headerFooterSnapshot == null ||
                !TryCreateHeaderFooterTextChrome(headerFooterSnapshot, pageNumber, pageCount, headerFooterDateTime, out HeaderFooterTextChrome chrome)) {
                return content;
            }

            string headerFooterSource = Name + "!headerFooter";
            var combinedDiagnostics = new List<OfficeImageExportDiagnostic>(content.Diagnostics.Count + 2);
            combinedDiagnostics.AddRange(content.Diagnostics);
            if (chrome.HasFormatting) {
                combinedDiagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Info,
                    ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation,
                    "Worksheet header/footer text formatting was rendered through the dependency-free image approximation path.",
                    headerFooterSource));
            }

            AddHeaderFooterImageDiagnostics(chrome, headerFooterSource, combinedDiagnostics);
            AddHeaderFooterFontDiagnostics(
                chrome,
                headerFooterSource,
                options.Fonts,
                combinedDiagnostics);

            double scale = rasterState.Scale;
            int headerHeight = chrome.HasHeader ? ResolveHeaderFooterBandHeight(chrome.HeaderImageHeightPoints, scale) : 0;
            int footerHeight = chrome.HasFooter ? ResolveHeaderFooterBandHeight(chrome.FooterImageHeightPoints, scale) : 0;
            bool pageSetupCanvasApplied = ShouldApplyPageSetupCanvas(GetPageSetup());
            int width = Math.Max(1, content.Width);
            int height = pageSetupCanvasApplied
                ? Math.Max(1, content.Height)
                : Math.Max(1, content.Height + headerHeight + footerHeight);
            double contentScaleRatio = 1D;
            OfficeRasterExportPlan? plan = null;
            if (format != OfficeImageExportFormat.Svg) {
                plan = OfficeRasterExportPlanner.Resolve(
                    width / scale,
                    height / scale,
                    rasterPlanningFormat,
                    options,
                    headerFooterSource);
                if (plan.Value.Limit.Scale < scale) {
                    contentScaleRatio = plan.Value.Limit.Scale / scale;
                    rasterState = ExcelRasterRenderState.FromPlan(plan.Value);
                    scale = rasterState.Scale;
                    headerHeight = chrome.HasHeader ? ResolveHeaderFooterBandHeight(chrome.HeaderImageHeightPoints, scale) : 0;
                    footerHeight = chrome.HasFooter ? ResolveHeaderFooterBandHeight(chrome.FooterImageHeightPoints, scale) : 0;
                    width = plan.Value.Limit.PixelWidth;
                    height = plan.Value.Limit.PixelHeight;
                    if (plan.Value.Diagnostic != null) {
                        combinedDiagnostics.Add(plan.Value.Diagnostic);
                    }
                }
            }

            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, combinedDiagnostics, headerFooterSource);
            IReadOnlyList<OfficeImageExportDiagnostic> diagnostics = combinedDiagnostics.Count == content.Diagnostics.Count
                ? content.Diagnostics
                : combinedDiagnostics.AsReadOnly();
            int contentY = pageSetupCanvasApplied ? 0 : headerHeight;
            int contentWidth = Math.Min(
                width,
                Math.Max(1, (int)Math.Ceiling(content.Width * contentScaleRatio)));
            int contentHeight = Math.Min(
                height,
                Math.Max(1, (int)Math.Ceiling(content.Height * contentScaleRatio)));
            if (!pageSetupCanvasApplied && contentScaleRatio < 1D) {
                contentHeight = ReconcileHeaderFooterRasterHeights(
                    height,
                    ref headerHeight,
                    ref footerHeight);
                contentY = headerHeight;
            }

            if (format == OfficeImageExportFormat.Svg) {
                OfficeImageLayer layer = OfficeImageLayer.FromSvgInner(
                    OfficeSvgFormatting.ExtractSvgInner(Encoding.UTF8.GetString(content.Bytes)),
                    0D,
                    contentY,
                    contentWidth,
                    contentHeight);
                return new OfficeImageExportResult(
                    format,
                    width,
                    height,
                    OfficeImageComposer.ComposeSvgBytes(
                        width,
                        height,
                        options.BackgroundColor,
                        new[] { layer },
                        beforeLayers: pageSetupCanvasApplied ? null : builder => AppendHeaderFooterSvgText(builder, chrome, width, height, headerHeight, scale, fallbackCodec),
                        afterLayers: pageSetupCanvasApplied ? builder => AppendHeaderFooterSvgText(builder, chrome, width, height, headerHeight, scale, fallbackCodec) : null),
                    content.Name,
                    content.Source,
                    diagnostics);
            }

            if (!OfficeRasterImageDecoder.TryDecode(content.Bytes, out OfficeRasterImage? contentImage) || contentImage == null) {
                return content;
            }

            OfficeImageLayer contentLayer = OfficeImageLayer.FromRaster(
                contentImage,
                0D,
                contentY,
                contentWidth,
                contentHeight);
            OfficeRasterImage image = OfficeImageComposer.ComposeRaster(
                width,
                height,
                options.BackgroundColor,
                new[] { contentLayer },
                beforeLayers: pageSetupCanvasApplied ? null : canvas => DrawHeaderFooterRaster(canvas, chrome, width, height, headerHeight, footerHeight, scale, fallbackCodec),
                afterLayers: pageSetupCanvasApplied ? canvas => DrawHeaderFooterRaster(canvas, chrome, width, height, headerHeight, footerHeight, scale, fallbackCodec) : null);
            return new OfficeImageExportResult(
                format,
                width,
                height,
                OfficeRasterImageEncoder.Encode(
                    image,
                    format,
                    rasterState.EncodingOptions),
                content.Name,
                content.Source,
                diagnostics);
        }

        private static int ReconcileHeaderFooterRasterHeights(
            int canvasHeight,
            ref int headerHeight,
            ref int footerHeight) {
            int maximumBandHeight = Math.Max(0, canvasHeight - 1);
            int totalBandHeight = checked(headerHeight + footerHeight);
            if (totalBandHeight > maximumBandHeight) {
                if (totalBandHeight == 0 || maximumBandHeight == 0) {
                    headerHeight = 0;
                    footerHeight = 0;
                } else {
                    headerHeight = Math.Min(
                        maximumBandHeight,
                        (int)Math.Round(
                            maximumBandHeight * (headerHeight / (double)totalBandHeight),
                            MidpointRounding.AwayFromZero));
                    footerHeight = maximumBandHeight - headerHeight;
                }
            }

            return canvasHeight - headerHeight - footerHeight;
        }

        private bool CanRenderHeaderFooterTextChrome(DateTime headerFooterDateTime) {
            if (!HasHeaderFooterContent()) {
                return true;
            }

            HeaderFooterSnapshot snapshot = GetHeaderFooter();

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

        private static void AddHeaderFooterFontDiagnostics(
            HeaderFooterTextChrome chrome,
            string source,
            OfficeFontFaceCollection fonts,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var reported = new HashSet<string>(StringComparer.Ordinal);
            AddHeaderFooterFontDiagnostics(chrome.HeaderLeft, source, fonts, diagnostics, reported);
            AddHeaderFooterFontDiagnostics(chrome.HeaderCenter, source, fonts, diagnostics, reported);
            AddHeaderFooterFontDiagnostics(chrome.HeaderRight, source, fonts, diagnostics, reported);
            AddHeaderFooterFontDiagnostics(chrome.FooterLeft, source, fonts, diagnostics, reported);
            AddHeaderFooterFontDiagnostics(chrome.FooterCenter, source, fonts, diagnostics, reported);
            AddHeaderFooterFontDiagnostics(chrome.FooterRight, source, fonts, diagnostics, reported);
        }

        private static void AddHeaderFooterFontDiagnostics(
            HeaderFooterTextSection section,
            string source,
            OfficeFontFaceCollection fonts,
            List<OfficeImageExportDiagnostic> diagnostics,
            HashSet<string> reported) {
            for (int index = 0; index < section.Runs.Count; index++) {
                HeaderFooterTextRun run = section.Runs[index];
                OfficeFontStyle style =
                    (run.Bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular) |
                    (run.Italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
                OfficeImageExportDiagnostic? diagnostic =
                    fonts.CreateSubstitutionDiagnostic(
                        run.Text,
                        run.FontFamily,
                        style,
                        source);
                if (diagnostic != null && reported.Add(diagnostic.Message)) {
                    diagnostics.Add(diagnostic);
                }
            }
        }

        private bool TryCreateHeaderFooterTextChrome(HeaderFooterSnapshot snapshot, int pageNumber, int pageCount, DateTime headerFooterDateTime, out HeaderFooterTextChrome chrome) {
            chrome = default;

            HeaderFooterVariantText selected = SelectHeaderFooterVariantText(snapshot, pageNumber);
            if (!TryCreateResolvedHeaderFooterTextChrome(
                selected.HeaderLeft,
                selected.HeaderCenter,
                selected.HeaderRight,
                selected.FooterLeft,
                selected.FooterCenter,
                selected.FooterRight,
                pageNumber,
                pageCount,
                headerFooterDateTime,
                out HeaderFooterTextChrome textChrome)) {
                return false;
            }

            chrome = textChrome.WithImages(
                SelectHeaderFooterImage(snapshot.HeaderLeftImage, selected.HeaderLeft),
                SelectHeaderFooterImage(snapshot.HeaderCenterImage, selected.HeaderCenter),
                SelectHeaderFooterImage(snapshot.HeaderRightImage, selected.HeaderRight),
                SelectHeaderFooterImage(snapshot.FooterLeftImage, selected.FooterLeft),
                SelectHeaderFooterImage(snapshot.FooterCenterImage, selected.FooterCenter),
                SelectHeaderFooterImage(snapshot.FooterRightImage, selected.FooterRight));
            return chrome.HasAnyContent;
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
                ResolveHeaderFooterFontFamily(),
                headerLeft,
                headerCenter,
                headerRight,
                footerLeft,
                footerCenter,
                footerRight);
            return true;
        }

        private string ResolveHeaderFooterFontFamily() {
            string familyName = GetWorkbookDefaultFontInfo()?.FamilyName ?? OfficeFontInfo.Default.FamilyName;
            return string.IsNullOrWhiteSpace(familyName)
                ? HeaderFooterFallbackFontFamily
                : familyName + ", " + HeaderFooterFallbackFontFamily;
        }

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
            double scale,
            IOfficeRasterImageCodec imageCodec) {
            double fontSize = HeaderFooterFontSize * scale;
            double padding = HeaderFooterHorizontalPadding * scale;
            OfficeTextZoneLayout zones = OfficeTextZoneLayout.CreateThreeColumn(width, padding, HeaderFooterZoneGap * scale);
            if (chrome.HasHeader && headerHeight > 0) {
                double y = Math.Max(0D, (headerHeight - fontSize) / 2D);
                DrawHeaderFooterRasterImages(canvas, chrome, isHeader: true, 0D, headerHeight, zones, scale, imageCodec);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderLeft, zones.Left, y, fontSize, chrome.FontFamily, OfficeTextAlignment.Left);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderCenter, zones.Center, y, fontSize, chrome.FontFamily, OfficeTextAlignment.Center);
                DrawHeaderFooterRasterLine(canvas, chrome.HeaderRight, zones.Right, y, fontSize, chrome.FontFamily, OfficeTextAlignment.Right);
            }

            if (chrome.HasFooter && footerHeight > 0) {
                double footerTop = height - footerHeight;
                double y = footerTop + Math.Max(0D, (footerHeight - fontSize) / 2D);
                DrawHeaderFooterRasterImages(canvas, chrome, isHeader: false, footerTop, footerHeight, zones, scale, imageCodec);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterLeft, zones.Left, y, fontSize, chrome.FontFamily, OfficeTextAlignment.Left);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterCenter, zones.Center, y, fontSize, chrome.FontFamily, OfficeTextAlignment.Center);
                DrawHeaderFooterRasterLine(canvas, chrome.FooterRight, zones.Right, y, fontSize, chrome.FontFamily, OfficeTextAlignment.Right);
            }
        }

        private static void DrawHeaderFooterRasterLine(
            OfficeRasterCanvas canvas,
            HeaderFooterTextSection section,
            OfficeTextZone zone,
            double y,
            double fontSize,
            string fontFamily,
            OfficeTextAlignment alignment) {
            if (!section.HasText) {
                return;
            }

            using (canvas.PushClipRectangle(zone.X, 0D, zone.Width, canvas.Height)) {
                if (section.HasFormatting) {
                    DrawHeaderFooterRasterRichLine(canvas, section, zone, y, fontSize, fontFamily, alignment);
                } else {
                    string displayText = ResolveHeaderFooterZoneText(section.Text, fontSize, zone.Width, (text, size) => canvas.MeasureText(text, size, fontFamily), alignment);
                    if (!string.IsNullOrWhiteSpace(displayText)) {
                        canvas.DrawTextLine(displayText, zone.AnchorX, y, fontSize, HeaderFooterTextColor, alignment: alignment, fontFamily: fontFamily);
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
            string fontFamily,
            OfficeTextAlignment alignment) {
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                section.ToOfficeRuns(fontSize, HeaderFooterTextColor, fontFamily),
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
            double scale,
            IOfficeRasterImageCodec imageCodec) {
            double fontSize = HeaderFooterFontSize * scale;
            double padding = HeaderFooterHorizontalPadding * scale;
            double lineHeight = fontSize * 1.2D;
            OfficeTextZoneLayout zones = OfficeTextZoneLayout.CreateThreeColumn(width, padding, HeaderFooterZoneGap * scale);
            OfficeTextMeasurer textMeasurer = OfficeTextMeasurer.Create(new OfficeFontInfo(chrome.FontFamily, fontSize));
            double MeasureText(string? text, double size, string? family) =>
                MeasureHeaderFooterSvgText(textMeasurer, text, size, string.IsNullOrWhiteSpace(family) ? chrome.FontFamily : family);
            if (chrome.HasHeader) {
                double baseline = Math.Max(fontSize, (headerHeight + fontSize) / 2D);
                AppendHeaderFooterSvgImages(builder, chrome, isHeader: true, 0D, headerHeight, zones, scale, imageCodec);
                AppendHeaderFooterSvgLine(builder, chrome.HeaderLeft, zones.Left, 0D, headerHeight, baseline, lineHeight, fontSize, chrome.FontFamily, OfficeTextAlignment.Left, "header-left", MeasureText);
                AppendHeaderFooterSvgLine(builder, chrome.HeaderCenter, zones.Center, 0D, headerHeight, baseline, lineHeight, fontSize, chrome.FontFamily, OfficeTextAlignment.Center, "header-center", MeasureText);
                AppendHeaderFooterSvgLine(builder, chrome.HeaderRight, zones.Right, 0D, headerHeight, baseline, lineHeight, fontSize, chrome.FontFamily, OfficeTextAlignment.Right, "header-right", MeasureText);
            }

            if (chrome.HasFooter) {
                int footerHeight = ResolveHeaderFooterBandHeight(chrome.FooterImageHeightPoints, scale);
                double footerTop = height - footerHeight;
                double baseline = footerTop + Math.Max(fontSize, (footerHeight + fontSize) / 2D);
                AppendHeaderFooterSvgImages(builder, chrome, isHeader: false, footerTop, footerHeight, zones, scale, imageCodec);
                AppendHeaderFooterSvgLine(builder, chrome.FooterLeft, zones.Left, footerTop, height - footerTop, baseline, lineHeight, fontSize, chrome.FontFamily, OfficeTextAlignment.Left, "footer-left", MeasureText);
                AppendHeaderFooterSvgLine(builder, chrome.FooterCenter, zones.Center, footerTop, height - footerTop, baseline, lineHeight, fontSize, chrome.FontFamily, OfficeTextAlignment.Center, "footer-center", MeasureText);
                AppendHeaderFooterSvgLine(builder, chrome.FooterRight, zones.Right, footerTop, height - footerTop, baseline, lineHeight, fontSize, chrome.FontFamily, OfficeTextAlignment.Right, "footer-right", MeasureText);
            }
        }

        private static double MeasureHeaderFooterSvgText(OfficeTextMeasurer measurer, string? text, double fontSize, string? fontFamily) {
            OfficeTextMeasurementStyle style = measurer.CreateStyle(new OfficeFontInfo(fontFamily, fontSize));
            return measurer.MeasureWidth(text, style);
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
            string fontFamily,
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
                AppendHeaderFooterSvgRichLine(builder, section, zone, baseline, fontSize, fontFamily, lineHeight, alignment, measure);
            } else {
                if (!string.IsNullOrWhiteSpace(section.Text)) {
                    builder.AppendSvgTextElement(
                        section.Text,
                        zone.AnchorX,
                        baseline,
                        lineHeight,
                        HeaderFooterTextColor,
                        fontFamily,
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
            string fontFamily,
            double lineHeight,
            OfficeTextAlignment alignment,
            Func<string?, double, string?, double> measure) {
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                section.ToOfficeRuns(fontSize, HeaderFooterTextColor, fontFamily),
                zone.Width,
                Math.Max(lineHeight, section.GetMaxResolvedFontSize(fontSize) * 1.2D),
                1.2D,
                measure,
                wrap: false,
                shrinkToFit: false,
                minimumFontSize: 1D,
                OfficeTextOverflowBehavior.Clip);
            if (layout.Lines.Count == 0) {
                return;
            }

            OfficeRichTextLine line = layout.Lines[0];
            double top = baseline - Math.Max(0D, (layout.LineHeight - Math.Max(1D, line.FontSize)) / 2D) - (Math.Max(1D, line.FontSize) * 0.84D);
            builder.AppendSvgRichTextBlock(
                layout,
                zone.X,
                top,
                zone.Width,
                layout.Height,
                alignment);
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
                string fontFamily,
                HeaderFooterTextSection headerLeft,
                HeaderFooterTextSection headerCenter,
                HeaderFooterTextSection headerRight,
                HeaderFooterTextSection footerLeft,
                HeaderFooterTextSection footerCenter,
                HeaderFooterTextSection footerRight) {
                FontFamily = fontFamily;
                HeaderLeft = headerLeft;
                HeaderCenter = headerCenter;
                HeaderRight = headerRight;
                FooterLeft = footerLeft;
                FooterCenter = footerCenter;
                FooterRight = footerRight;
            }

            private HeaderFooterTextChrome(
                string fontFamily,
                HeaderFooterTextSection headerLeft,
                HeaderFooterTextSection headerCenter,
                HeaderFooterTextSection headerRight,
                HeaderFooterTextSection footerLeft,
                HeaderFooterTextSection footerCenter,
                HeaderFooterTextSection footerRight,
                HeaderFooterImageSnapshot? headerLeftImage,
                HeaderFooterImageSnapshot? headerCenterImage,
                HeaderFooterImageSnapshot? headerRightImage,
                HeaderFooterImageSnapshot? footerLeftImage,
                HeaderFooterImageSnapshot? footerCenterImage,
                HeaderFooterImageSnapshot? footerRightImage)
                : this(fontFamily, headerLeft, headerCenter, headerRight, footerLeft, footerCenter, footerRight) {
                HeaderLeftImage = headerLeftImage;
                HeaderCenterImage = headerCenterImage;
                HeaderRightImage = headerRightImage;
                FooterLeftImage = footerLeftImage;
                FooterCenterImage = footerCenterImage;
                FooterRightImage = footerRightImage;
            }

            internal string FontFamily { get; }
            internal HeaderFooterTextSection HeaderLeft { get; }
            internal HeaderFooterTextSection HeaderCenter { get; }
            internal HeaderFooterTextSection HeaderRight { get; }
            internal HeaderFooterTextSection FooterLeft { get; }
            internal HeaderFooterTextSection FooterCenter { get; }
            internal HeaderFooterTextSection FooterRight { get; }
            internal HeaderFooterImageSnapshot? HeaderLeftImage { get; }
            internal HeaderFooterImageSnapshot? HeaderCenterImage { get; }
            internal HeaderFooterImageSnapshot? HeaderRightImage { get; }
            internal HeaderFooterImageSnapshot? FooterLeftImage { get; }
            internal HeaderFooterImageSnapshot? FooterCenterImage { get; }
            internal HeaderFooterImageSnapshot? FooterRightImage { get; }
            internal bool HasHeader => HeaderLeft.HasText || HeaderCenter.HasText || HeaderRight.HasText || HeaderLeftImage != null || HeaderCenterImage != null || HeaderRightImage != null;
            internal bool HasFooter => FooterLeft.HasText || FooterCenter.HasText || FooterRight.HasText || FooterLeftImage != null || FooterCenterImage != null || FooterRightImage != null;
            internal bool HasAnyContent => HasHeader || HasFooter;
            internal double HeaderImageHeightPoints => MaxImageHeight(HeaderLeftImage, HeaderCenterImage, HeaderRightImage);
            internal double FooterImageHeightPoints => MaxImageHeight(FooterLeftImage, FooterCenterImage, FooterRightImage);
            internal bool HasFormatting =>
                HeaderLeft.HasFormatting ||
                HeaderCenter.HasFormatting ||
                HeaderRight.HasFormatting ||
                FooterLeft.HasFormatting ||
                FooterCenter.HasFormatting ||
                FooterRight.HasFormatting;

            internal HeaderFooterTextChrome WithImages(
                HeaderFooterImageSnapshot? headerLeft,
                HeaderFooterImageSnapshot? headerCenter,
                HeaderFooterImageSnapshot? headerRight,
                HeaderFooterImageSnapshot? footerLeft,
                HeaderFooterImageSnapshot? footerCenter,
                HeaderFooterImageSnapshot? footerRight) =>
                new HeaderFooterTextChrome(
                    FontFamily,
                    HeaderLeft,
                    HeaderCenter,
                    HeaderRight,
                    FooterLeft,
                    FooterCenter,
                    FooterRight,
                    headerLeft,
                    headerCenter,
                    headerRight,
                    footerLeft,
                    footerCenter,
                    footerRight);

            private static double MaxImageHeight(params HeaderFooterImageSnapshot?[] images) {
                double height = 0D;
                for (int index = 0; index < images.Length; index++) {
                    if (images[index] != null) {
                        height = Math.Max(height, images[index]!.HeightPoints);
                    }
                }

                return height;
            }
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
