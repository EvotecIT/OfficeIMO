using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    internal static class GoogleSlidesBatchCompiler {
        internal static GoogleSlidesBatch Build(PowerPointPresentation presentation, GoogleSlidesSaveOptions options) {
            var report = new TranslationReport();
            var plan = new GoogleSlidesTranslationPlan(report) { SlideCount = presentation.Slides.Count };
            string? title = !string.IsNullOrWhiteSpace(options.Title) ? options.Title! : presentation.BuiltinDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title)) title = "Presentation";
            var batch = new GoogleSlidesBatch(title!, presentation.SlideSize.WidthPoints, presentation.SlideSize.HeightPoints, plan);

            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                PowerPointSlide source = presentation.Slides[slideIndex];
                var target = new GoogleSlidesSlide(ObjectId("slide", slideIndex, 0), slideIndex);
                PowerPointSlideBackground background = source.GetBackground();
                if (background.Kind == PowerPointSlideBackgroundKind.SolidColor) target.BackgroundColorHex = background.Color;
                if (source.Notes.TryGetExistingText(out string notes)) { target.SpeakerNotes = notes; plan.SpeakerNotesCount++; }

                PowerPointShape[] unsupported = source.Shapes.Where(IsUnsupported).ToArray();
                if (unsupported.Length > 0 && options.ComplexSlides == GoogleSlidesComplexSlideMode.RasterizeComplexSlides) {
                    byte[] bytes = source.ToPng(new PowerPointImageExportOptions { IncludeSlideBackground = true, IncludeHiddenShapes = false });
                    target.IsRasterized = true;
                    target.Add(new GoogleSlidesImage(ObjectId("render", slideIndex, 0), 0, 0, batch.WidthPoints, batch.HeightPoints, bytes, "image/png", $"slide-{slideIndex + 1}.png"));
                    plan.RasterizedSlideCount++;
                    plan.UnsupportedElementCount += unsupported.Length;
                    report.Add(TranslationSeverity.Warning, "ComplexSlides", $"Slide {slideIndex + 1} contains {unsupported.Length} element(s) without a dependable native Slides equivalent and was rendered to PNG.",
                        path: $"slide/{slideIndex + 1}", code: "SLIDES.COMPLEX_SLIDE.RASTERIZED", action: TranslationAction.Rasterize, count: unsupported.Length);
                    batch.Add(target);
                    continue;
                }

                int elementIndex = 0;
                foreach (PowerPointShape shape in source.Shapes.OrderBy(shape => shape.DrawingOrder)) {
                    string id = ObjectId("element", slideIndex, elementIndex++);
                    switch (shape) {
                        case PowerPointTextBox textBox:
                            var text = new GoogleSlidesTextBox(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints, textBox.Text);
                            PowerPointTextRun? firstRun = textBox.Paragraphs.SelectMany(paragraph => paragraph.Runs).FirstOrDefault();
                            if (firstRun != null) {
                                text.Bold = firstRun.Bold; text.Italic = firstRun.Italic; text.Underline = firstRun.Underline;
                                text.FontSize = firstRun.FontSize; text.FontFamily = firstRun.FontName; text.ForegroundColorHex = firstRun.Color;
                                text.Hyperlink = firstRun.Hyperlink?.AbsoluteUri;
                            }
                            target.Add(text); plan.NativeTextBoxCount++;
                            break;
                        case PowerPointTable table:
                            IReadOnlyList<IReadOnlyList<string>> cells = table.RowItems.Select(row => (IReadOnlyList<string>)row.Cells.Select(cell => cell.Text).ToArray()).ToArray();
                            target.Add(new GoogleSlidesTable(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints, cells));
                            plan.NativeTableCount++;
                            break;
                        case PowerPointPicture picture:
                            target.Add(new GoogleSlidesImage(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints,
                                picture.GetImageBytes(), picture.ContentType ?? "image/png", $"picture-{slideIndex + 1}-{elementIndex}.png"));
                            plan.NativeImageCount++;
                            break;
                        case PowerPointAutoShape autoShape:
                            target.Add(new GoogleSlidesShape(id, shape.LeftPoints, shape.TopPoints, shape.WidthPoints, shape.HeightPoints, MapShape(autoShape)));
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

        private static bool IsUnsupported(PowerPointShape shape) => shape.ShapeContentType == PowerPointShapeContentType.Chart
            || shape.ShapeContentType == PowerPointShapeContentType.SmartArt
            || shape.ShapeContentType == PowerPointShapeContentType.Media
            || shape.ShapeContentType == PowerPointShapeContentType.Group
            || shape.ShapeContentType == PowerPointShapeContentType.Connector
            || shape.ShapeContentType == PowerPointShapeContentType.OleObject
            || shape.ShapeContentType == PowerPointShapeContentType.Unknown;

        private static string MapShape(PowerPointAutoShape shape) {
            string name = shape.ShapeType?.ToString() ?? string.Empty;
            if (name.IndexOf("Ellipse", StringComparison.OrdinalIgnoreCase) >= 0) return "ELLIPSE";
            if (name.IndexOf("Triangle", StringComparison.OrdinalIgnoreCase) >= 0) return "TRIANGLE";
            if (name.IndexOf("Arrow", StringComparison.OrdinalIgnoreCase) >= 0) return "RIGHT_ARROW";
            return "RECTANGLE";
        }

        private static string ObjectId(string kind, int slideIndex, int elementIndex) => $"officeimo_{kind}_{slideIndex + 1:D4}_{elementIndex + 1:D4}";
    }
}
