using System.Globalization;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>Inspects generated or imported slides against a structured accessibility policy.</summary>
        public OfficeAccessibilityReport InspectAccessibility(PowerPointAccessibilityOptions? options = null) {
            ThrowIfDisposed();
            PowerPointAccessibilityOptions resolved = (options ??
                PowerPointAccessibilityOptions.ForProfile(OfficeAccessibilityPolicyProfile.Default))
                .CloneValidated();
            var findings = new List<OfficeAccessibilityFinding>();
            if (resolved.RequireDocumentTitle && string.IsNullOrWhiteSpace(Properties.Title)) {
                findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error,
                    "accessibility-missing-document-title",
                    "The presentation package does not define a document title.", "Document"));
            }
            InspectAccessibilitySlides(resolved, findings);
            return new OfficeAccessibilityReport("PowerPoint", resolved.Profile, findings);
        }

        /// <summary>Lists each inspected slide's detected title and its shapes in reading order with their accessibility metadata.</summary>
        /// <remarks>Hidden slides and the shape-tree limits follow <paramref name="options"/>; the policy checks are not run.</remarks>
        public IReadOnlyList<PowerPointAccessibilitySlideInfo> InspectReadingOrder(PowerPointAccessibilityOptions? options = null) {
            ThrowIfDisposed();
            PowerPointAccessibilityOptions resolved = (options ?? new PowerPointAccessibilityOptions()).CloneValidated();
            return InspectAccessibilitySlides(resolved, findings: null);
        }

        private List<PowerPointAccessibilitySlideInfo> InspectAccessibilitySlides(PowerPointAccessibilityOptions resolved,
            List<OfficeAccessibilityFinding>? findings) {
            var slideInfos = new List<PowerPointAccessibilitySlideInfo>(_slides.Count);
            for (int slideIndex = 0; slideIndex < _slides.Count; slideIndex++) {
                PowerPointSlide slide = _slides[slideIndex];
                if (slide.Hidden && !resolved.IncludeHiddenSlides) continue;
                List<PowerPointShape> shapes = EnumerateAccessibilityShapes(slide, resolved).ToList();
                string? slideTitle = FindSlideTitle(
                    slide.EnumerateShapesDeep(slide.Shapes, includeHidden: false,
                        resolved.MaximumShapeCount, resolved.MaximumGroupDepth),
                    SlideSize.Height.ToPoints());
                if (findings != null && resolved.RequireSlideTitles && string.IsNullOrWhiteSpace(slideTitle)) {
                    findings.Add(SlideFinding(OfficeAccessibilitySeverity.Error,
                        "accessibility-missing-slide-title", "The slide does not expose a recognizable title.", slideIndex));
                }

                var shapeInfos = new List<PowerPointAccessibilityShapeInfo>(shapes.Count);
                foreach (PowerPointShape shape in shapes) {
                    shapeInfos.Add(new PowerPointAccessibilityShapeInfo(shape.ReadingOrder, shape.Id, shape.Name,
                        shape.Title, shape.AltText, shape.Decorative, shape.Language, shape.ShapeContentType));
                    if (findings != null) InspectShapeAccessibility(slide, slideIndex, shape, shapes, resolved, findings);
                }

                slideInfos.Add(new PowerPointAccessibilitySlideInfo(slideIndex, slideTitle, shapeInfos));
            }
            return slideInfos;
        }

        private static OfficeAccessibilityFinding SlideFinding(OfficeAccessibilitySeverity severity, string code, string message,
            int slideIndex, PowerPointShape? shape = null, bool isHeuristic = false,
            double? measuredValue = null, double? requiredValue = null) =>
            new(severity, code, message, "Slide " + (slideIndex + 1).ToString(CultureInfo.InvariantCulture), slideIndex,
                shape?.Id, shape?.Name, isHeuristic, measuredValue, requiredValue);

        private static IEnumerable<PowerPointShape> EnumerateAccessibilityShapes(PowerPointSlide slide,
            PowerPointAccessibilityOptions options) {
            return slide.EnumerateShapesDeep(slide.GetInheritedShapesForExport().Concat(slide.Shapes),
                includeHidden: false, options.MaximumShapeCount, options.MaximumGroupDepth);
        }

        private static string? FindSlideTitle(IEnumerable<PowerPointShape> shapes, double slideHeightPoints) {
            List<PowerPointTextBox> textBoxes = shapes.OfType<PowerPointTextBox>()
                .Where(box => !string.IsNullOrWhiteSpace(box.Text)).ToList();
            PowerPointTextBox? title = textBoxes.FirstOrDefault(box =>
                box.ShapePlaceholderType == PowerPointPlaceholderType.Title ||
                box.ShapePlaceholderType == PowerPointPlaceholderType.CenteredTitle);
            title ??= textBoxes.FirstOrDefault(box =>
                string.Equals(box.Title, "Slide title", StringComparison.OrdinalIgnoreCase));
            title ??= textBoxes.Where(box => box.Top.ToPoints() <= slideHeightPoints * 0.32D)
                .OrderByDescending(box => box.FontSize?.ToPoints() ?? 0).ThenBy(box => box.Top.ToPoints()).FirstOrDefault();
            return string.IsNullOrWhiteSpace(title?.Text) ? null : NormalizeVisibleText(title!.Text);
        }

        private static void InspectShapeAccessibility(PowerPointSlide slide, int slideIndex, PowerPointShape shape,
            IReadOnlyList<PowerPointShape> shapes,
            PowerPointAccessibilityOptions options, IList<OfficeAccessibilityFinding> findings) {
            if (shape.Decorative) return;
            bool informativeVisual = IsInformativeVisual(shape);
            if (informativeVisual && options.RequireShapeTitles && string.IsNullOrWhiteSpace(shape.Title)) {
                findings.Add(ShapeFinding(options.Profile, "accessibility-missing-shape-title",
                    "The informative visual does not define a concise accessibility title.", slideIndex, shape));
            }
            if (informativeVisual && options.RequireAlternativeText && string.IsNullOrWhiteSpace(shape.AltText)) {
                findings.Add(SlideFinding(OfficeAccessibilitySeverity.Error,
                    "accessibility-missing-alternative-text",
                    "The informative visual is neither decorative nor described with alternative text.",
                    slideIndex, shape));
            }
            if (options.RequireLanguage && HasVisibleText(shape) && string.IsNullOrWhiteSpace(shape.Language)) {
                findings.Add(ShapeFinding(options.Profile, "accessibility-missing-language",
                    "The text-bearing shape does not define an explicit language tag.", slideIndex, shape));
            }
            if (options.RequireTableHeaders && shape is PowerPointTable table && !table.HeaderRow) {
                findings.Add(SlideFinding(OfficeAccessibilitySeverity.Error,
                    "accessibility-missing-table-header",
                    "The table does not mark its first row as a header row.", slideIndex, shape));
            }
            if (options.CheckMeaningfulLinks && shape is PowerPointTextBox textBox) {
                InspectLinks(textBox, slideIndex, shape, options, findings);
            }
            if (options.CheckContrast && shape is PowerPointTextBox contrastTextBox) {
                InspectContrast(slide, contrastTextBox, shapes, slideIndex, options, findings);
            }
            if (options.CheckColorOnlyMeaning && shape is PowerPointChart chart) {
                InspectChartMeaning(chart, slideIndex, options, findings);
            }
        }

        private static void InspectLinks(PowerPointTextBox textBox, int slideIndex, PowerPointShape shape,
            PowerPointAccessibilityOptions options, IList<OfficeAccessibilityFinding> findings) {
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
                int index = 0;
                while (index < runs.Count) {
                    PowerPointTextRun run = runs[index];
                    Uri? hyperlink = run.Hyperlink;
                    if (hyperlink == null) {
                        index++;
                        continue;
                    }

                    int end = index + 1;
                    while (end < runs.Count && SameHyperlink(hyperlink, runs[end].Hyperlink)) end++;
                    string label = string.Concat(runs.Skip(index).Take(end - index).Select(item => item.Text));
                    if (!PowerPointTextRun.IsMeaningfulLinkLabel(label)) {
                        findings.Add(SlideFinding(OfficeAccessibilitySeverity.Error,
                            "accessibility-unclear-link-label",
                            "Hyperlink text must describe its destination without surrounding context.",
                            slideIndex, shape));
                    }

                    if (options.Profile == OfficeAccessibilityPolicyProfile.Strict) {
                        for (int runIndex = index; runIndex < end; runIndex++) {
                            if (string.IsNullOrWhiteSpace(runs[runIndex].HyperlinkTooltip)) {
                                findings.Add(SlideFinding(OfficeAccessibilitySeverity.Warning,
                                    "accessibility-missing-link-tooltip",
                                    "Strict policy recommends an accessible hyperlink tooltip.",
                                    slideIndex, shape));
                            }
                        }
                    }
                    index = end;
                }
            }
        }

        private static bool SameHyperlink(Uri expected, Uri? candidate) =>
            candidate != null && string.Equals(expected.OriginalString, candidate.OriginalString,
                StringComparison.Ordinal);

        private static void InspectContrast(PowerPointSlide slide, PowerPointTextBox textBox,
            IReadOnlyList<PowerPointShape> shapes, int slideIndex,
            PowerPointAccessibilityOptions options, IList<OfficeAccessibilityFinding> findings) {
            OfficeColor? background = ResolveBackgroundColor(slide, textBox, shapes);
            if (!background.HasValue) return;
            bool inspectedRun = false;
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                foreach (PowerPointParagraphInline inline in paragraph.InlineNodes) {
                    PowerPointTextRun? run = inline.Run;
                    if (run == null || string.IsNullOrWhiteSpace(run.Text)) continue;
                    OfficeColor? foreground = PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(run, paragraph, textBox);
                    double? fontSize = run.FontSize?.ToPoints() ?? textBox.FontSize?.ToPoints();
                    if (!foreground.HasValue || !fontSize.HasValue) continue;
                    inspectedRun = true;
                    bool large = fontSize.Value >= options.LargeTextThreshold.ToPoints() ||
                        run.Bold && fontSize.Value >= options.LargeBoldTextThreshold.ToPoints();
                    double required = large ? options.MinimumLargeTextContrastRatio : options.MinimumTextContrastRatio;
                    double measured = OfficeColorContrast.ContrastRatio(foreground.Value, background.Value);
                    if (measured + 0.000001D < required) {
                        findings.Add(SlideFinding(OfficeAccessibilitySeverity.Error,
                            "accessibility-low-contrast",
                            "Text contrast is " + measured.ToString("0.##", CultureInfo.InvariantCulture) +
                            ":1; the selected policy requires at least " +
                            required.ToString("0.##", CultureInfo.InvariantCulture) + ":1.",
                            slideIndex, textBox, measuredValue: measured, requiredValue: required));
                        return;
                    }
                }
            }

            if (!inspectedRun && !string.IsNullOrWhiteSpace(textBox.Text)) {
                OfficeColor? foreground = PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(null, textBox.Paragraphs.FirstOrDefault(), textBox);
                double? fontSize = textBox.FontSize?.ToPoints();
                if (foreground.HasValue && fontSize.HasValue) {
                    double required = fontSize.Value >= options.LargeTextThreshold.ToPoints()
                        ? options.MinimumLargeTextContrastRatio : options.MinimumTextContrastRatio;
                    double measured = OfficeColorContrast.ContrastRatio(foreground.Value, background.Value);
                    if (measured + 0.000001D < required) {
                        findings.Add(SlideFinding(OfficeAccessibilitySeverity.Error,
                            "accessibility-low-contrast", "Text does not meet the selected contrast policy.",
                            slideIndex, textBox, measuredValue: measured, requiredValue: required));
                    }
                }
            }
        }

        private static void InspectChartMeaning(PowerPointChart chart, int slideIndex,
            PowerPointAccessibilityOptions options, IList<OfficeAccessibilityFinding> findings) {
            if (!chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) || snapshot.Data.Series.Count < 2) return;
            string description = chart.AltText ?? string.Empty;
            if (description.IndexOf("Data summary:", StringComparison.OrdinalIgnoreCase) >= 0) return;
            findings.Add(SlideFinding(
                options.Profile == OfficeAccessibilityPolicyProfile.Strict
                    ? OfficeAccessibilitySeverity.Error : OfficeAccessibilitySeverity.Warning,
                "accessibility-chart-color-only-meaning",
                "The multi-series chart lacks a data summary, so series meaning may depend on color alone.",
                slideIndex, chart));
        }

        private static bool IsInformativeVisual(PowerPointShape shape) =>
            shape.ShapeContentType == PowerPointShapeContentType.Picture ||
            shape.ShapeContentType == PowerPointShapeContentType.Media ||
            shape.ShapeContentType == PowerPointShapeContentType.Chart ||
            shape.ShapeContentType == PowerPointShapeContentType.Table ||
            shape.ShapeContentType == PowerPointShapeContentType.SmartArt;

        private static bool HasVisibleText(PowerPointShape shape) {
            if (shape is PowerPointTextBox textBox) return !string.IsNullOrWhiteSpace(textBox.Text);
            if (shape is PowerPointTable table) {
                for (int row = 0; row < table.RowCount; row++)
                    for (int column = 0; column < table.ColumnCount; column++)
                        if (!string.IsNullOrWhiteSpace(table.GetCell(row, column).Text)) return true;
            }
            return false;
        }

        private static OfficeAccessibilityFinding ShapeFinding(OfficeAccessibilityPolicyProfile profile,
            string code, string message, int slideIndex, PowerPointShape shape) =>
            SlideFinding(profile == OfficeAccessibilityPolicyProfile.Strict
                    ? OfficeAccessibilitySeverity.Error : OfficeAccessibilitySeverity.Warning,
                code, message, slideIndex, shape);

        private static OfficeColor? ResolveBackgroundColor(PowerPointSlide slide, PowerPointTextBox textBox,
            IReadOnlyList<PowerPointShape> shapes) {
            if (textBox.ResolveFillColorForExport() is OfficeColor ownFill && (textBox.FillTransparency ?? 0) < 50) {
                return ownFill;
            }
            int textBoxIndex = IndexOfReference(shapes, textBox);
            IEnumerable<PowerPointShape> earlierShapes = textBoxIndex >= 0
                ? shapes.Take(textBoxIndex)
                : shapes.Where(shape => shape.DrawingOrder < textBox.DrawingOrder);
            PowerPointShape? containingShape = earlierShapes
                .Where(shape => !ReferenceEquals(shape, textBox) && !shape.Hidden &&
                                shape.ResolveFillColorForExport().HasValue &&
                                (shape.FillTransparency ?? 0) < 50 && Contains(shape, textBox))
                .LastOrDefault();
            if (containingShape != null) return containingShape.ResolveFillColorForExport();
            PowerPointSlideBackground background = slide.GetBackground();
            return background.Kind == PowerPointSlideBackgroundKind.SolidColor
                ? ParseColor(background.Color) : null;
        }

        private static int IndexOfReference(IReadOnlyList<PowerPointShape> shapes, PowerPointShape target) {
            for (int i = 0; i < shapes.Count; i++) {
                if (ReferenceEquals(shapes[i], target)) return i;
            }
            return -1;
        }

        private static bool Contains(PowerPointShape outer, PowerPointShape inner) =>
            inner.Left.ToPoints() >= outer.Left.ToPoints() - 0.5D && inner.Top.ToPoints() >= outer.Top.ToPoints() - 0.5D &&
            inner.Left.ToPoints() + inner.Width.ToPoints() <= outer.Left.ToPoints() + outer.Width.ToPoints() + 0.5D &&
            inner.Top.ToPoints() + inner.Height.ToPoints() <= outer.Top.ToPoints() + outer.Height.ToPoints() + 0.5D;

        private static OfficeColor? ParseColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            string normalized = value!.Trim().TrimStart('#');
            return OfficeColor.TryParseHex(normalized, out OfficeColor color) ? color : (OfficeColor?)null;
        }

        private static string NormalizeVisibleText(string value) =>
            string.Join(" ", (value ?? string.Empty).Split(new[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries));
    }
}
