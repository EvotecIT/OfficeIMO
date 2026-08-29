using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    internal static class GoogleSlidesBatchCompiler {
        internal static GoogleSlidesTranslationPlan BuildPlan(PowerPointPresentation presentation, GoogleSlidesSaveOptions options) =>
            Build(presentation, options, materializeRasterImages: false).Plan;

        internal static GoogleSlidesBatch Build(
            PowerPointPresentation presentation,
            GoogleSlidesSaveOptions options,
            bool materializeRasterImages = true) {
            var report = new TranslationReport();
            var plan = new GoogleSlidesTranslationPlan(report) { SlideCount = presentation.Slides.Count };
            string? title = !string.IsNullOrWhiteSpace(options.Title) ? options.Title! : presentation.BuiltinDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title)) title = "Presentation";
            var batch = new GoogleSlidesBatch(title!, presentation.SlideSize.WidthPoints, presentation.SlideSize.HeightPoints, plan);

            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                PowerPointSlide source = presentation.Slides[slideIndex];
                var target = new GoogleSlidesSlide(ObjectId("slide", slideIndex, 0), slideIndex) {
                    IsSkipped = source.Hidden,
                };
                PowerPointSlideBackground background = source.GetBackground();
                if (background.Kind == PowerPointSlideBackgroundKind.SolidColor) target.BackgroundColorHex = NormalizeColorHex(background.Color);
                else if (IsSupportedSlidesBackgroundImage(background)) {
                    target.BackgroundImage = new GoogleSlidesImage(
                        ObjectId("background", slideIndex, 0),
                        0,
                        0,
                        batch.WidthPoints,
                        batch.HeightPoints,
                        background.ImageBytes!,
                        background.ImageContentType!,
                        $"background-{slideIndex + 1}{ImageExtension(background.ImageContentType)}");
                }
                if (source.Notes.TryGetExistingText(out string notes)) { target.SpeakerNotes = notes; plan.SpeakerNotesCount++; }

                PowerPointShape[] visibleShapes = source.Shapes.Where(shape => !shape.Hidden).ToArray();
                PowerPointShape[] unsupported = visibleShapes.Where(IsUnsupported).ToArray();
                bool unsupportedBackground = IsUnsupportedBackground(background);
                int unsupportedFeatureCount = unsupported.Length + (unsupportedBackground ? 1 : 0);
                if (unsupportedFeatureCount > 0 && options.ComplexSlides == GoogleSlidesComplexSlideMode.RasterizeComplexSlides) {
                    target.IsRasterized = true;
                    target.BackgroundColorHex = null;
                    target.BackgroundImage = null;
                    if (materializeRasterImages) {
                        byte[] bytes = source.ToPng(new PowerPointImageExportOptions { IncludeSlideBackground = true, IncludeHiddenShapes = false });
                        target.Add(new GoogleSlidesImage(ObjectId("render", slideIndex, 0), 0, 0, batch.WidthPoints, batch.HeightPoints, bytes, "image/png", $"slide-{slideIndex + 1}.png"));
                    }
                    plan.RasterizedSlideCount++;
                    plan.UnsupportedElementCount += unsupportedFeatureCount;
                    string rasterMessage = materializeRasterImages
                        ? $"Slide {slideIndex + 1} contains {unsupportedFeatureCount} feature(s) without a dependable native Slides equivalent and was rendered to PNG."
                        : $"Slide {slideIndex + 1} contains {unsupportedFeatureCount} feature(s) without a dependable native Slides equivalent and will be rendered to PNG during export.";
                    report.Add(TranslationSeverity.Warning, "ComplexSlides", rasterMessage,
                        path: $"slide/{slideIndex + 1}", code: "SLIDES.COMPLEX_SLIDE.RASTERIZED", action: TranslationAction.Rasterize, count: unsupportedFeatureCount);
                    batch.Add(target);
                    continue;
                }

                if (unsupportedBackground) {
                    plan.UnsupportedElementCount++;
                    report.Add(
                        TranslationSeverity.Warning,
                        "Backgrounds",
                        UnsupportedBackgroundMessage(background),
                        path: $"slide/{slideIndex + 1}/background",
                        code: "SLIDES.BACKGROUND.SKIPPED",
                        action: TranslationAction.Skip);
                }

                int elementIndex = 0;
                foreach (PowerPointShape shape in visibleShapes.OrderBy(shape => shape.DrawingOrder)) {
                    string id = ObjectId("element", slideIndex, elementIndex++);
                    switch (shape) {
                        case PowerPointTextBox textBox:
                            var text = PreserveTransform(new GoogleSlidesTextBox(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints, BuildTextContent(textBox)), shape);
                            if (!textBox.UsesTextBoxGeometry && TryMapShape(textBox.ShapeType, out string textShapeType)) text.ShapeType = textShapeType;
                            PreserveShapeStyle(text.Style, shape);
                            PopulateTextRuns(text, textBox);
                            GoogleSlidesTextStyleRun? firstRun = text.TextRuns.FirstOrDefault();
                            if (firstRun != null) {
                                text.Bold = firstRun.Bold; text.Italic = firstRun.Italic; text.Underline = firstRun.Underline;
                                text.Strikethrough = firstRun.Strikethrough;
                                text.SmallCaps = firstRun.SmallCaps;
                                text.BaselineOffset = firstRun.BaselineOffset;
                                text.FontSize = firstRun.FontSize; text.FontFamily = firstRun.FontFamily; text.ForegroundColorHex = firstRun.ForegroundColorHex;
                                text.Hyperlink = firstRun.Hyperlink;
                            }
                            target.Add(text); plan.NativeTextBoxCount++;
                            break;
                        case PowerPointTable table when !HasMergedCells(table):
                            IReadOnlyList<IReadOnlyList<GoogleSlidesTableCell>> cells = table.RowItems
                                .Select(row => (IReadOnlyList<GoogleSlidesTableCell>)row.Cells
                                    .Select(BuildTableCell).ToArray()).ToArray();
                            target.Add(PreserveTransform(new GoogleSlidesTable(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints, cells), shape));
                            plan.NativeTableCount++;
                            break;
                        case PowerPointPicture picture when IsSupportedSlidesImage(picture):
                            target.Add(PreserveTransform(new GoogleSlidesImage(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints,
                                picture.GetImageBytes(), picture.ContentType ?? "image/png", $"picture-{slideIndex + 1}-{elementIndex}{ImageExtension(picture.ContentType)}"), shape));
                            plan.NativeImageCount++;
                            break;
                        case PowerPointPicture picture when HasImageCrop(picture):
                            plan.UnsupportedElementCount++;
                            report.Add(
                                TranslationSeverity.Warning,
                                "Images",
                                $"Skipped cropped image '{picture.Name ?? id}' because Google Slides exposes image crop properties as read-only. Use RasterizeComplexSlides to preserve its rendered appearance.",
                                path: $"slide/{slideIndex + 1}/{picture.Name ?? id}",
                                code: "SLIDES.IMAGE.CROP_SKIPPED",
                                action: TranslationAction.Skip);
                            break;
                        case PowerPointPicture picture:
                            plan.UnsupportedElementCount++;
                            report.Add(
                                TranslationSeverity.Warning,
                                "Images",
                                $"Skipped image '{picture.Name ?? id}' with content type '{picture.ContentType ?? "unknown"}' because Google Slides createImage accepts PNG, JPEG, or GIF only.",
                                path: $"slide/{slideIndex + 1}/{picture.Name ?? id}",
                                code: "SLIDES.IMAGE.FORMAT_SKIPPED",
                                action: TranslationAction.Skip);
                            break;
                        case PowerPointAutoShape autoShape when TryMapShape(autoShape, out string slidesShapeType):
                            GoogleSlidesShape slidesShape = PreserveTransform(new GoogleSlidesShape(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints, slidesShapeType), shape);
                            PreserveShapeStyle(slidesShape.Style, shape);
                            target.Add(slidesShape);
                            plan.NativeShapeCount++;
                            break;
                        default:
                            plan.UnsupportedElementCount++;
                            report.Add(TranslationSeverity.Warning, "PageElements", $"Skipped {shape.ShapeContentType} element '{shape.Name ?? id}' because PreferNativeAndReport was selected.",
                                path: $"slide/{slideIndex + 1}/{shape.Name ?? id}", code: "SLIDES.PAGE_ELEMENT.SKIPPED", action: TranslationAction.Skip);
                            break;
                    }
                }
                batch.Add(target);
            }
            if (presentation.Slides.Count == 0) report.Add(TranslationSeverity.Warning, "Slides", "The source presentation contains no slides.", code: "SLIDES.EMPTY_SOURCE", action: TranslationAction.Skip);
            return batch;
        }

        private static string BuildTextContent(PowerPointTextBox textBox) =>
            BuildTextContent(textBox.Paragraphs, textBox.TextBody?.ListStyle, textBox.MasterTextStyle);

        private static string BuildTextContent(
            IReadOnlyList<PowerPointParagraph> paragraphs,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) => string.Join(
                "\n",
                paragraphs.Select(paragraph => string.Concat(paragraph.InlineNodes.Select(node =>
                    GetGoogleInlineText(node, paragraph, listStyle, masterTextStyle)))));

        private static void PopulateTextRuns(GoogleSlidesTextBox target, PowerPointTextBox source) {
            PopulateTextRuns(target.TextRuns, source.Paragraphs, source.TextBody?.ListStyle, source.MasterTextStyle);
        }

        private static GoogleSlidesTableCell BuildTableCell(PowerPointTableCell cell) {
            var textRuns = new List<GoogleSlidesTextStyleRun>();
            A.TextBody? textBody = cell.Cell.TextBody;
            OpenXmlCompositeElement? masterTextStyle = cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?
                .SlideMaster?.TextStyles?.OtherStyle;
            PopulateTextRuns(textRuns, cell.Paragraphs, textBody?.ListStyle, masterTextStyle);
            return new GoogleSlidesTableCell(BuildTextContent(cell.Paragraphs, textBody?.ListStyle, masterTextStyle), textRuns);
        }

        private static void PopulateTextRuns(
            ICollection<GoogleSlidesTextStyleRun> target,
            IReadOnlyList<PowerPointParagraph> paragraphs,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) {
            int offset = 0;
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
                PowerPointParagraph paragraph = paragraphs[paragraphIndex];
                foreach (PowerPointParagraphInline node in paragraph.InlineNodes) {
                    if (node.Kind == PowerPointParagraphInlineKind.LineBreak) {
                        offset++;
                        continue;
                    }
                    PowerPointTextRun? run = node.Run;
                    if (run == null) continue;
                    EffectiveGoogleRunStyle effective = ResolveEffectiveRunStyle(run, paragraph, listStyle, masterTextStyle);
                    string text = GetGoogleText(run, effective);
                    int endIndex = offset + text.Length;
                    if (endIndex > offset) {
                        target.Add(new GoogleSlidesTextStyleRun {
                            StartIndex = offset,
                            EndIndex = endIndex,
                            Bold = effective.Bold ?? false,
                            Italic = effective.Italic ?? false,
                            Underline = effective.Underline ?? false,
                            Strikethrough = effective.Strikethrough ?? false,
                            SmallCaps = effective.Capitalization == PowerPointCapitalization.SmallCaps,
                            BaselineOffset = ToGoogleBaselineOffset(effective.BaselinePercent),
                            FontSize = effective.FontSizePoints.HasValue
                                ? (int?)Math.Round(effective.FontSizePoints.Value, MidpointRounding.AwayFromZero)
                                : null,
                            FontFamily = effective.FontName,
                            ForegroundColorHex = NormalizeColorHex(effective.Color),
                            Hyperlink = run.Hyperlink?.AbsoluteUri,
                        });
                    }
                    offset = endIndex;
                }
                if (paragraphIndex + 1 < paragraphs.Count) offset++;
            }
        }

        internal static string GetGoogleInlineText(
            PowerPointParagraphInline node,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) {
            if (node.Kind == PowerPointParagraphInlineKind.LineBreak) return "\n";
            if (node.Run == null) return node.Text;
            return GetGoogleText(node.Run, ResolveEffectiveRunStyle(node.Run, paragraph, listStyle, masterTextStyle));
        }

        internal static string GetGoogleText(PowerPointTextRun run) {
            string text = run.Text ?? string.Empty;
            if (run.Capitalization != PowerPointCapitalization.AllCaps) return text;
            return text.ToUpper(ResolveRunCulture(run.Language));
        }

        internal static string GetGoogleText(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) =>
            GetGoogleText(run, ResolveEffectiveRunStyle(run, paragraph, listStyle, masterTextStyle));

        internal static bool IsGoogleSmallCaps(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) =>
            ResolveEffectiveRunStyle(run, paragraph, listStyle, masterTextStyle).Capitalization == PowerPointCapitalization.SmallCaps;

        internal static string? GetGoogleBaselineOffset(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) =>
            ToGoogleBaselineOffset(ResolveEffectiveRunStyle(run, paragraph, listStyle, masterTextStyle).BaselinePercent);

        private static string GetGoogleText(PowerPointTextRun run, EffectiveGoogleRunStyle effective) {
            string text = run.Text ?? string.Empty;
            if (effective.Capitalization != PowerPointCapitalization.AllCaps) return text;
            return text.ToUpper(ResolveRunCulture(effective.Language));
        }

        internal static EffectiveGoogleRunStyle ResolveEffectiveRunStyle(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) {
            PowerPointEffectiveRunStyle effective = PowerPointEffectiveRunStyleResolver.Resolve(
                run, paragraph, listStyle, masterTextStyle);
            return new EffectiveGoogleRunStyle(
                effective.Capitalization,
                effective.BaselinePercent,
                effective.Language,
                effective.FontSizePoints,
                effective.FontName,
                effective.Color,
                effective.Bold,
                effective.Italic,
                effective.UnderlineStyle.HasValue
                    ? effective.UnderlineStyle.Value != PowerPointUnderlineStyle.None
                    : (bool?)null,
                effective.StrikeStyle.HasValue
                    ? effective.StrikeStyle.Value != PowerPointStrikeStyle.None
                    : (bool?)null);
        }

        internal readonly struct EffectiveGoogleRunStyle {
            internal EffectiveGoogleRunStyle(
                PowerPointCapitalization? capitalization,
                double? baselinePercent,
                string? language,
                double? fontSizePoints,
                string? fontName,
                string? color,
                bool? bold,
                bool? italic,
                bool? underline,
                bool? strikethrough) {
                Capitalization = capitalization;
                BaselinePercent = baselinePercent;
                Language = language;
                FontSizePoints = fontSizePoints;
                FontName = fontName;
                Color = color;
                Bold = bold;
                Italic = italic;
                Underline = underline;
                Strikethrough = strikethrough;
            }

            internal PowerPointCapitalization? Capitalization { get; }
            internal double? BaselinePercent { get; }
            internal string? Language { get; }
            internal double? FontSizePoints { get; }
            internal string? FontName { get; }
            internal string? Color { get; }
            internal bool? Bold { get; }
            internal bool? Italic { get; }
            internal bool? Underline { get; }
            internal bool? Strikethrough { get; }
        }

        private static CultureInfo ResolveRunCulture(string? language) {
            if (string.IsNullOrWhiteSpace(language)) return CultureInfo.InvariantCulture;
            try {
                return CultureInfo.GetCultureInfo(language);
            } catch (CultureNotFoundException) {
                return CultureInfo.InvariantCulture;
            }
        }

        internal static string? ToGoogleBaselineOffset(double? baselinePercent) => baselinePercent switch {
            > 0 => "SUPERSCRIPT",
            < 0 => "SUBSCRIPT",
            _ => null,
        };

        private static bool IsUnsupported(PowerPointShape shape) => (shape is PowerPointPicture picture && (!IsSupportedSlidesImage(picture) || HasImageCrop(picture)))
            || (shape is PowerPointAutoShape autoShape && !TryMapShape(autoShape, out _))
            || (shape is PowerPointTable table && HasMergedCells(table))
            || shape.ShapeContentType == PowerPointShapeContentType.Chart
            || shape.ShapeContentType == PowerPointShapeContentType.SmartArt
            || shape.ShapeContentType == PowerPointShapeContentType.Media
            || shape.ShapeContentType == PowerPointShapeContentType.Group
            || shape.ShapeContentType == PowerPointShapeContentType.Connector
            || shape.ShapeContentType == PowerPointShapeContentType.OleObject
            || shape.ShapeContentType == PowerPointShapeContentType.Unknown;

        private static bool IsSupportedSlidesImage(PowerPointPicture picture) {
            return IsSupportedSlidesImageContentType(picture.ContentType) && !HasImageCrop(picture);
        }

        private static bool HasImageCrop(PowerPointPicture picture) =>
            Math.Abs(picture.CropLeftRatio) > double.Epsilon
            || Math.Abs(picture.CropTopRatio) > double.Epsilon
            || Math.Abs(picture.CropRightRatio) > double.Epsilon
            || Math.Abs(picture.CropBottomRatio) > double.Epsilon;

        private static bool IsSupportedSlidesImageContentType(string? imageContentType) {
            string contentType = imageContentType ?? string.Empty;
            return contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase)
                || contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase)
                || contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)
                || contentType.Equals("image/gif", StringComparison.OrdinalIgnoreCase);
        }

        private static T PreserveTransform<T>(T element, PowerPointShape source) where T : GoogleSlidesElement {
            element.RotationDegrees = source.Rotation ?? 0d;
            element.HorizontalFlip = source.HorizontalFlip == true;
            element.VerticalFlip = source.VerticalFlip == true;
            return element;
        }

        private static void PreserveShapeStyle(GoogleSlidesShapeStyle target, PowerPointShape source) {
            target.FillColorHex = NormalizeColorHex(source.FillColor);
            target.FillTransparencyPercent = source.FillTransparency;
            target.OutlineColorHex = NormalizeColorHex(source.OutlineColor);
            target.OutlineWidthPoints = source.OutlineWidthPoints;
        }

        private static bool IsSupportedSlidesBackgroundImage(PowerPointSlideBackground background) =>
            background.Kind == PowerPointSlideBackgroundKind.Image
            && background.ImageBytes is { Length: > 0 }
            && !background.HasImageCrop
            && IsSupportedSlidesImageContentType(background.ImageContentType);

        private static bool IsUnsupportedBackground(PowerPointSlideBackground background) =>
            (background.Kind == PowerPointSlideBackgroundKind.Image && !IsSupportedSlidesBackgroundImage(background))
            || background.Kind == PowerPointSlideBackgroundKind.LinearGradient
            || background.Kind == PowerPointSlideBackgroundKind.Unsupported;

        private static string UnsupportedBackgroundMessage(PowerPointSlideBackground background) => background.Kind switch {
            PowerPointSlideBackgroundKind.Image when background.HasImageCrop => "Skipped the cropped slide image background because Google Slides stretched-picture backgrounds cannot preserve PowerPoint source cropping.",
            PowerPointSlideBackgroundKind.Image => $"Skipped the slide image background with content type '{background.ImageContentType ?? "unknown"}' because Google Slides stretched-picture backgrounds accept PNG, JPEG, or GIF content.",
            PowerPointSlideBackgroundKind.LinearGradient => "Skipped the slide gradient background because Google Slides page backgrounds support solid fills but not PowerPoint gradient fills.",
            _ => $"Skipped the slide background because it has no dependable native Google Slides equivalent{(string.IsNullOrWhiteSpace(background.UnsupportedReason) ? "." : $": {background.UnsupportedReason}")}",
        };

        private static string ImageExtension(string? contentType) => (contentType ?? string.Empty).ToLowerInvariant() switch {
            "image/jpeg" or "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            _ => ".png",
        };

        internal static string? NormalizeColorHex(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            string candidate = value!.Trim().TrimStart('#');
            if (candidate.Length >= 6) candidate = candidate.Substring(candidate.Length - 6);
            if (candidate.Length != 6) return null;
            for (int index = 0; index < candidate.Length; index++) {
                char character = candidate[index];
                if (!((character >= '0' && character <= '9')
                    || (character >= 'a' && character <= 'f')
                    || (character >= 'A' && character <= 'F'))) {
                    return null;
                }
            }
            return candidate.ToUpperInvariant();
        }

        private static bool HasMergedCells(PowerPointTable table) {
            return table.RowItems.SelectMany(row => row.Cells).Any(cell => cell.IsMergedCell || cell.IsMergeAnchor);
        }

        private static bool TryMapShape(PowerPointAutoShape shape, out string slidesShapeType) =>
            TryMapShape(shape.ShapeType, out slidesShapeType);

        private static bool TryMapShape(OfficePresetShapeType? shapeType, out string slidesShapeType) {
            if (shapeType == OfficePresetShapeType.Rectangle) slidesShapeType = "RECTANGLE";
            else if (shapeType == OfficePresetShapeType.RoundRectangle) slidesShapeType = "ROUND_RECTANGLE";
            else if (shapeType == OfficePresetShapeType.Ellipse) slidesShapeType = "ELLIPSE";
            else if (shapeType == OfficePresetShapeType.Triangle) slidesShapeType = "TRIANGLE";
            else if (shapeType == OfficePresetShapeType.RightTriangle) slidesShapeType = "RIGHT_TRIANGLE";
            else if (shapeType == OfficePresetShapeType.Parallelogram) slidesShapeType = "PARALLELOGRAM";
            else if (shapeType == OfficePresetShapeType.Trapezoid) slidesShapeType = "TRAPEZOID";
            else if (shapeType == OfficePresetShapeType.Diamond) slidesShapeType = "DIAMOND";
            else if (shapeType == OfficePresetShapeType.RightArrow) slidesShapeType = "RIGHT_ARROW";
            else {
                slidesShapeType = string.Empty;
                return false;
            }

            return true;
        }

        private static string ObjectId(string kind, int slideIndex, int elementIndex) => $"officeimo_{kind}_{slideIndex + 1:D4}_{elementIndex + 1:D4}";
    }
}
