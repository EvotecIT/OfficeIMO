using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>Inspects generated or imported slides against a structured accessibility policy.</summary>
        public PowerPointAccessibilityReport InspectAccessibility(PowerPointAccessibilityOptions? options = null) {
            ThrowIfDisposed();
            PowerPointAccessibilityOptions resolved = (options ??
                PowerPointAccessibilityOptions.ForProfile(PowerPointAccessibilityPolicyProfile.Default))
                .CloneValidated();
            var findings = new List<PowerPointAccessibilityFinding>();
            var slideInfos = new List<PowerPointAccessibilitySlideInfo>(_slides.Count);

            if (resolved.RequireDocumentTitle && string.IsNullOrWhiteSpace(BuiltinDocumentProperties.Title)) {
                findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Error,
                    "Accessibility.MissingDocumentTitle",
                    "The presentation package does not define a document title."));
            }

            for (int slideIndex = 0; slideIndex < _slides.Count; slideIndex++) {
                PowerPointSlide slide = _slides[slideIndex];
                if (slide.Hidden && !resolved.IncludeHiddenSlides) continue;
                List<PowerPointShape> shapes = EnumerateAccessibilityShapes(slide, resolved).ToList();
                string? slideTitle = FindSlideTitle(shapes, SlideSize.HeightPoints);
                if (resolved.RequireSlideTitles && string.IsNullOrWhiteSpace(slideTitle)) {
                    findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Error,
                        "Accessibility.MissingSlideTitle", "The slide does not expose a recognizable title.", slideIndex));
                }

                var shapeInfos = new List<PowerPointAccessibilityShapeInfo>(shapes.Count);
                for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++) {
                    PowerPointShape shape = shapes[shapeIndex];
                    shapeInfos.Add(new PowerPointAccessibilityShapeInfo(shape.ReadingOrder, shape.Id, shape.Name,
                        shape.Title, shape.Description, shape.Decorative, shape.Language, shape.ShapeContentType));
                    InspectShapeAccessibility(slide, slideIndex, shape, shapes, resolved, findings);
                }

                slideInfos.Add(new PowerPointAccessibilitySlideInfo(slideIndex, slideTitle, shapeInfos));
            }

            return new PowerPointAccessibilityReport(resolved.Profile, slideInfos, findings);
        }

        private static IEnumerable<PowerPointShape> EnumerateAccessibilityShapes(PowerPointSlide slide,
            PowerPointAccessibilityOptions options) {
            return slide.EnumerateShapesDeep(slide.GetInheritedShapesForExport().Concat(slide.Shapes),
                includeHidden: false, options.MaximumShapeCount, options.MaximumGroupDepth);
        }

        private static string? FindSlideTitle(IEnumerable<PowerPointShape> shapes, double slideHeightPoints) {
            List<PowerPointTextBox> textBoxes = shapes.OfType<PowerPointTextBox>()
                .Where(box => !string.IsNullOrWhiteSpace(box.Text)).ToList();
            PowerPointTextBox? title = textBoxes.FirstOrDefault(box =>
                box.ShapePlaceholderType == PlaceholderValues.Title ||
                box.ShapePlaceholderType == PlaceholderValues.CenteredTitle);
            title ??= textBoxes.FirstOrDefault(box =>
                string.Equals(box.Title, "Slide title", StringComparison.OrdinalIgnoreCase));
            title ??= textBoxes.Where(box => box.TopPoints <= slideHeightPoints * 0.32D)
                .OrderByDescending(box => box.FontSize ?? 0).ThenBy(box => box.TopPoints).FirstOrDefault();
            return string.IsNullOrWhiteSpace(title?.Text) ? null : NormalizeVisibleText(title!.Text);
        }

        private static void InspectShapeAccessibility(PowerPointSlide slide, int slideIndex, PowerPointShape shape,
            IReadOnlyList<PowerPointShape> shapes,
            PowerPointAccessibilityOptions options, IList<PowerPointAccessibilityFinding> findings) {
            if (shape.Decorative) return;
            bool informativeVisual = IsInformativeVisual(shape);
            if (informativeVisual && options.RequireShapeTitles && string.IsNullOrWhiteSpace(shape.Title)) {
                findings.Add(ShapeFinding(options.Profile, "Accessibility.MissingShapeTitle",
                    "The informative visual does not define a concise accessibility title.", slideIndex, shape));
            }
            if (informativeVisual && options.RequireAlternativeText && string.IsNullOrWhiteSpace(shape.Description)) {
                findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Error,
                    "Accessibility.MissingAlternativeText",
                    "The informative visual is neither decorative nor described with alternative text.",
                    slideIndex, shape.Id, shape.Name));
            }
            if (options.RequireLanguage && HasVisibleText(shape) && string.IsNullOrWhiteSpace(shape.Language)) {
                findings.Add(ShapeFinding(options.Profile, "Accessibility.MissingLanguage",
                    "The text-bearing shape does not define an explicit language tag.", slideIndex, shape));
            }
            if (options.RequireTableHeaders && shape is PowerPointTable table && !table.HeaderRow) {
                findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Error,
                    "Accessibility.MissingTableHeader",
                    "The table does not mark its first row as a header row.", slideIndex, shape.Id, shape.Name));
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
            PowerPointAccessibilityOptions options, IList<PowerPointAccessibilityFinding> findings) {
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
                        findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Error,
                            "Accessibility.UnclearLinkLabel",
                            "Hyperlink text must describe its destination without surrounding context.",
                            slideIndex, shape.Id, shape.Name));
                    }

                    if (options.Profile == PowerPointAccessibilityPolicyProfile.Strict) {
                        for (int runIndex = index; runIndex < end; runIndex++) {
                            if (string.IsNullOrWhiteSpace(runs[runIndex].HyperlinkTooltip)) {
                                findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Warning,
                                    "Accessibility.MissingLinkTooltip",
                                    "Strict policy recommends an accessible hyperlink tooltip.",
                                    slideIndex, shape.Id, shape.Name));
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
            PowerPointAccessibilityOptions options, IList<PowerPointAccessibilityFinding> findings) {
            OfficeColor? background = ResolveBackgroundColor(slide, textBox, shapes);
            if (!background.HasValue) return;
            bool inspectedRun = false;
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                foreach (PowerPointTextRun run in paragraph.Runs) {
                    if (string.IsNullOrWhiteSpace(run.Text)) continue;
                    OfficeColor? foreground = ParseColor(run.Color ?? textBox.Color);
                    double? fontSize = run.FontSize ?? textBox.FontSize;
                    if (!foreground.HasValue || !fontSize.HasValue) continue;
                    inspectedRun = true;
                    bool large = fontSize.Value >= options.LargeTextThresholdPoints ||
                        run.Bold && fontSize.Value >= options.LargeBoldTextThresholdPoints;
                    double required = large ? options.MinimumLargeTextContrastRatio : options.MinimumTextContrastRatio;
                    double measured = OfficeColorContrast.ContrastRatio(foreground.Value, background.Value);
                    if (measured + 0.000001D < required) {
                        findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Error,
                            "Accessibility.LowContrast",
                            "Text contrast is " + measured.ToString("0.##", CultureInfo.InvariantCulture) +
                            ":1; the selected policy requires at least " +
                            required.ToString("0.##", CultureInfo.InvariantCulture) + ":1.",
                            slideIndex, textBox.Id, textBox.Name, measured, required));
                        return;
                    }
                }
            }

            if (!inspectedRun && !string.IsNullOrWhiteSpace(textBox.Text)) {
                OfficeColor? foreground = ParseColor(textBox.Color);
                double? fontSize = textBox.FontSize;
                if (foreground.HasValue && fontSize.HasValue) {
                    double required = fontSize.Value >= options.LargeTextThresholdPoints
                        ? options.MinimumLargeTextContrastRatio : options.MinimumTextContrastRatio;
                    double measured = OfficeColorContrast.ContrastRatio(foreground.Value, background.Value);
                    if (measured + 0.000001D < required) {
                        findings.Add(new PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity.Error,
                            "Accessibility.LowContrast", "Text does not meet the selected contrast policy.",
                            slideIndex, textBox.Id, textBox.Name, measured, required));
                    }
                }
            }
        }

        private static void InspectChartMeaning(PowerPointChart chart, int slideIndex,
            PowerPointAccessibilityOptions options, IList<PowerPointAccessibilityFinding> findings) {
            if (!chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) || snapshot.Data.Series.Count < 2) return;
            string description = chart.Description ?? string.Empty;
            if (description.IndexOf("Data summary:", StringComparison.OrdinalIgnoreCase) >= 0) return;
            findings.Add(new PowerPointAccessibilityFinding(
                options.Profile == PowerPointAccessibilityPolicyProfile.Strict
                    ? PowerPointAccessibilitySeverity.Error : PowerPointAccessibilitySeverity.Warning,
                "Accessibility.ChartColorOnlyMeaning",
                "The multi-series chart lacks a data summary, so series meaning may depend on color alone.",
                slideIndex, chart.Id, chart.Name));
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
                for (int row = 0; row < table.Rows; row++)
                    for (int column = 0; column < table.Columns; column++)
                        if (!string.IsNullOrWhiteSpace(table.GetCell(row, column).Text)) return true;
            }
            return false;
        }

        private static PowerPointAccessibilityFinding ShapeFinding(PowerPointAccessibilityPolicyProfile profile,
            string code, string message, int slideIndex, PowerPointShape shape) =>
            new(profile == PowerPointAccessibilityPolicyProfile.Strict
                    ? PowerPointAccessibilitySeverity.Error : PowerPointAccessibilitySeverity.Warning,
                code, message, slideIndex, shape.Id, shape.Name);

        private static OfficeColor? ResolveBackgroundColor(PowerPointSlide slide, PowerPointTextBox textBox,
            IReadOnlyList<PowerPointShape> shapes) {
            if (!string.IsNullOrWhiteSpace(textBox.FillColor) && (textBox.FillTransparency ?? 0) < 50) {
                return ParseColor(textBox.FillColor);
            }
            int textBoxIndex = IndexOfReference(shapes, textBox);
            IEnumerable<PowerPointShape> earlierShapes = textBoxIndex >= 0
                ? shapes.Take(textBoxIndex)
                : shapes.Where(shape => shape.DrawingOrder < textBox.DrawingOrder);
            PowerPointShape? containingShape = earlierShapes
                .Where(shape => !ReferenceEquals(shape, textBox) && !shape.Hidden &&
                                !string.IsNullOrWhiteSpace(shape.FillColor) &&
                                (shape.FillTransparency ?? 0) < 50 && Contains(shape, textBox))
                .LastOrDefault();
            if (containingShape != null) return ParseColor(containingShape.FillColor);
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
            inner.LeftPoints >= outer.LeftPoints - 0.5D && inner.TopPoints >= outer.TopPoints - 0.5D &&
            inner.LeftPoints + inner.WidthPoints <= outer.LeftPoints + outer.WidthPoints + 0.5D &&
            inner.TopPoints + inner.HeightPoints <= outer.TopPoints + outer.HeightPoints + 0.5D;

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
