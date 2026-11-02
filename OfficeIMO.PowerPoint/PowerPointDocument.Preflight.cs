using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        ///     Inspects authored slide bounds, measured text fit, peer collisions, image relationships, and
        ///     shared visual-snapshot diagnostics without changing the presentation.
        /// </summary>
        public PowerPointDeckPreflightReport InspectPreflight(PowerPointDeckPreflightOptions? options = null) {
            ThrowIfDisposed();
            PowerPointDeckPreflightOptions resolved = options?.Clone() ?? new PowerPointDeckPreflightOptions();
            var findings = new List<PowerPointDeckPreflightFinding>();
            long slideWidth = SlideSize.Width.ToEmus();
            long slideHeight = SlideSize.Height.ToEmus();
            int inspectedShapeCount = 0;

            for (int slideIndex = 0; slideIndex < Slides.Count; slideIndex++) {
                PowerPointSlide slide = Slides[slideIndex];
                InspectSlide(slide, slideIndex, slideWidth, slideHeight, resolved, findings,
                    ref inspectedShapeCount);
            }

            return new PowerPointDeckPreflightReport(Slides.Count, findings);
        }

        /// <summary>
        ///     Runs deck preflight and saves only when no finding meets the configured failure threshold.
        /// </summary>
        internal PowerPointDeckPreflightReport SaveWithPreflight(PowerPointDeckPreflightOptions? options = null) {
            PowerPointDeckPreflightOptions resolved = options?.Clone() ?? new PowerPointDeckPreflightOptions();
            PowerPointDeckPreflightReport report = InspectPreflight(resolved);
            if (report.HasFindingsAtOrAbove(resolved.FailureSeverity)) {
                _discardChangesOnDispose = true;
            }
            report.ThrowIfFindings(resolved.FailureSeverity);
            Save();
            return report;
        }

        internal PowerPointDeckPreflightReport Preflight(PowerPointDeckPreflightOptions? options = null) =>
            InspectPreflight(options);

        private static void InspectSlide(PowerPointSlide slide, int slideIndex, long slideWidth, long slideHeight,
            PowerPointDeckPreflightOptions options, IList<PowerPointDeckPreflightFinding> findings,
            ref int inspectedShapeCount) {
            InspectShapeTree(slide, slide.Shapes, slideIndex,
                new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0L),OfficeIMO.OfficeLength.Emus(0L),OfficeIMO.OfficeLength.Emus(slideWidth),OfficeIMO.OfficeLength.Emus(slideHeight)), options, findings, null,
                0, ref inspectedShapeCount);

            if (options.IncludeVisualSnapshotDiagnostics) {
                InspectVisualSnapshot(slide, slideIndex, findings);
            }
        }

        private static void InspectShapeTree(PowerPointSlide slide, IReadOnlyList<PowerPointShape> shapes,
            int slideIndex, PowerPointLayoutBox canvas, PowerPointDeckPreflightOptions options,
            IList<PowerPointDeckPreflightFinding> findings, int? containingShapeIndex, int groupDepth,
            ref int inspectedShapeCount) {
            if (groupDepth > options.MaximumGroupDepth) {
                throw new InvalidOperationException("The grouped-shape nesting depth exceeds the configured limit.");
            }
            for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++) {
                PowerPointShape shape = shapes[shapeIndex];
                if (++inspectedShapeCount > options.MaximumShapeCount) {
                    throw new InvalidOperationException("The shape count exceeds the configured inspection limit.");
                }
                int reportShapeIndex = containingShapeIndex ?? shapeIndex;
                if (shape.Hidden) {
                    continue;
                }

                if (options.DetectOffSlideShapes) {
                    InspectBounds(shape, reportShapeIndex, slideIndex, canvas, options, findings);
                }
                if ((options.DetectTextOverflow || options.DetectUnreadableFontReduction) &&
                    shape is PowerPointTextBox textBox) {
                    InspectText(textBox, reportShapeIndex, slideIndex, options, findings);
                }
                if (options.DetectMissingVisualAssets && shape is PowerPointPicture picture) {
                    InspectPicture(picture, reportShapeIndex, slideIndex, findings);
                }
                if (shape is PowerPointGroupShape groupShape) {
                    int remainingShapeCount = options.MaximumShapeCount - inspectedShapeCount;
                    IReadOnlyList<PowerPointShape> children = slide.GetGroupChildren(groupShape, remainingShapeCount);
                    InspectShapeTree(slide, children, slideIndex,
                        PowerPointSlide.GetGroupChildBounds(groupShape, children),
                        options, findings, reportShapeIndex, groupDepth + 1, ref inspectedShapeCount);
                }
            }

            if (options.DetectShapeCollisions) {
                InspectCollisions(shapes, slideIndex, options, findings, containingShapeIndex);
            }
        }

        private static bool InheritsGeometry(PowerPointShape shape) =>
            shape.Element is DocumentFormat.OpenXml.Presentation.Shape element &&
            element.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null &&
            element.ShapeProperties?.Transform2D == null;

        private static bool IsLineGeometry(PowerPointShape shape) {
            if (shape.Element is DocumentFormat.OpenXml.Presentation.ConnectionShape) return true;
            DocumentFormat.OpenXml.Drawing.PresetGeometry? geometry = shape.Element
                .Descendants<DocumentFormat.OpenXml.Drawing.PresetGeometry>().FirstOrDefault();
            DocumentFormat.OpenXml.Drawing.ShapeTypeValues? preset = geometry?.Preset?.Value;
            return preset == DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Line ||
                   preset == DocumentFormat.OpenXml.Drawing.ShapeTypeValues.StraightConnector1;
        }

        private static void InspectBounds(PowerPointShape shape, int shapeIndex, int slideIndex,
            PowerPointLayoutBox canvas, PowerPointDeckPreflightOptions options,
            IList<PowerPointDeckPreflightFinding> findings) {
            PowerPointLayoutBox bounds = shape.Bounds;
            // A placeholder without its own transform takes its geometry from the layout or master; it has no
            // slide-level bounds to check.
            if (bounds.Width.ToEmus() <= 0L && bounds.Height.ToEmus() <= 0L && InheritsGeometry(shape)) return;
            // A horizontal or vertical line has no extent across its direction; only a line without length is invalid.
            bool line = IsLineGeometry(shape);
            if (line ? bounds.Width.ToEmus() <= 0L && bounds.Height.ToEmus() <= 0L : bounds.Width.ToEmus() <= 0L || bounds.Height.ToEmus() <= 0L) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "layout-invalid-bounds",
                    "Shape has a non-positive width or height.", slideIndex, shapeIndex, shape, bounds));
                return;
            }

            if (bounds.Left.ToEmus() < canvas.Left.ToEmus() || bounds.Top.ToEmus() < canvas.Top.ToEmus() ||
                bounds.Right.ToEmus() > canvas.Right.ToEmus() || bounds.Bottom.ToEmus() > canvas.Bottom.ToEmus()) {
                long bleed = options.MaximumDecorativeBleed.ToEmus();
                bool allowedDecorativeBleed = options.AllowDecorativeShapeBleed &&
                    shape is PowerPointAutoShape &&
                    bounds.Left.ToEmus() >= canvas.Left.ToEmus() - bleed && bounds.Top.ToEmus() >= canvas.Top.ToEmus() - bleed &&
                    bounds.Right.ToEmus() <= canvas.Right.ToEmus() + bleed && bounds.Bottom.ToEmus() <= canvas.Bottom.ToEmus() + bleed &&
                    bounds.Right.ToEmus() > canvas.Left.ToEmus() && bounds.Bottom.ToEmus() > canvas.Top.ToEmus() &&
                    bounds.Left.ToEmus() < canvas.Right.ToEmus() && bounds.Top.ToEmus() < canvas.Bottom.ToEmus();
                if (allowedDecorativeBleed) {
                    return;
                }
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "layout-shape-off-slide",
                    "Shape extends beyond its slide or group canvas.", slideIndex, shapeIndex, shape, bounds));
            }
        }

        private static void InspectText(PowerPointTextBox textBox, int shapeIndex, int slideIndex,
            PowerPointDeckPreflightOptions options, IList<PowerPointDeckPreflightFinding> findings) {
            string text = textBox.Text;
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            if (textBox.Width.ToPoints() <= 0D || textBox.Height.ToPoints() <= 0D) return;
            double authoredFontSize = ResolveFontSize(textBox, options.DefaultFontSize.ToPoints());
            string fontName = ResolveFontName(textBox);
            var metrics = OfficeDrawingTextLayout.CreateMetrics(new OfficeDrawing(textBox.Width.ToPoints(), textBox.Height.ToPoints()));
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            // Inspect the same paragraph flow, native fitting and mixed-run metrics used
            // by image/PDF export instead of a second plain-text wrapping approximation.
            OfficeDrawing drawing = PowerPointSlideImageRenderer.CreateTextBoxDrawing(textBox,
                textBox.Width.ToPoints(), textBox.Height.ToPoints(), diagnostics, metrics.MeasureText, fontName);
            bool clipped = diagnostics.Any(diagnostic => diagnostic.Code is "powerpoint-text-overflow" or "powerpoint-text-layout-limit");
            double resolvedFontSize = authoredFontSize;
            if (textBox.TextAutoFit == PowerPointTextAutoFit.Normal) {
                resolvedFontSize = PowerPointSlideTextInspector.Read(drawing).Select(span => span.FontSize)
                    .DefaultIfEmpty(authoredFontSize).Max();
            }

            if (options.DetectTextOverflow && clipped &&
                textBox.TextAutoFit != PowerPointTextAutoFit.Shape) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "text-clipped",
                    "Measured text does not fit inside the authored text box.", slideIndex, shapeIndex, textBox,
                    textBox.Bounds, resolvedFontSize));
            }

            if (options.DetectUnreadableFontReduction && resolvedFontSize < options.MinimumReadableFontSize.ToPoints()) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Warning, "text-unreadable-font-reduction",
                    "Text fit resolves to " + resolvedFontSize.ToString("0.##", CultureInfo.InvariantCulture) +
                    " pt, below the configured " +
                    options.MinimumReadableFontSize.ToPoints().ToString("0.##", CultureInfo.InvariantCulture) +
                    " pt threshold.", slideIndex, shapeIndex, textBox, textBox.Bounds, resolvedFontSize));
            }
        }

        private static double ResolveFontSize(PowerPointTextBox textBox, double fallback) {
            double maximum = 0D;
            for (int paragraphIndex = 0; paragraphIndex < textBox.Paragraphs.Count; paragraphIndex++) {
                IReadOnlyList<PowerPointTextRun> runs = textBox.Paragraphs[paragraphIndex].Runs;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    maximum = Math.Max(maximum, runs[runIndex].FontSize?.ToPoints() ?? 0D);
                }
            }

            return maximum > 0D ? maximum : textBox.FontSize?.ToPoints() ?? fallback;
        }

        private static string ResolveFontName(PowerPointTextBox textBox) {
            for (int paragraphIndex = 0; paragraphIndex < textBox.Paragraphs.Count; paragraphIndex++) {
                IReadOnlyList<PowerPointTextRun> runs = textBox.Paragraphs[paragraphIndex].Runs;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    if (!string.IsNullOrWhiteSpace(runs[runIndex].FontName)) {
                        return PowerPointEffectiveRunStyleResolver.ResolveThemeFontReference(
                            runs[runIndex].FontName, runs[runIndex].OwnerPart) ?? "Aptos";
                    }
                }
            }

            return PowerPointEffectiveRunStyleResolver.ResolveThemeFontReference(textBox.FontName,
                textBox.RelationshipOwnerPart) ?? "Aptos";
        }

        private static void InspectPicture(PowerPointPicture picture, int shapeIndex, int slideIndex,
            IList<PowerPointDeckPreflightFinding> findings) {
            try {
                if (string.IsNullOrWhiteSpace(picture.ContentType)) {
                    findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "asset-missing-image-part",
                        "Picture does not resolve to an embedded image part.", slideIndex, shapeIndex, picture,
                        picture.Bounds));
                }
            } catch (Exception exception) when (exception is ArgumentOutOfRangeException ||
                                                exception is KeyNotFoundException ||
                                                exception is InvalidOperationException) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "asset-broken-image-relationship",
                    "Picture relationship cannot be resolved: " + exception.Message, slideIndex, shapeIndex,
                    picture, picture.Bounds));
            }
        }

        private static void InspectCollisions(IReadOnlyList<PowerPointShape> shapes, int slideIndex,
            PowerPointDeckPreflightOptions options, IList<PowerPointDeckPreflightFinding> findings,
            int? containingShapeIndex) {
            double tolerance = options.CollisionTolerance.ToEmus();
            for (int leftIndex = 0; leftIndex < shapes.Count; leftIndex++) {
                PowerPointShape left = shapes[leftIndex];
                if (ShouldSkipCollisionShape(left)) continue;

                for (int rightIndex = leftIndex + 1; rightIndex < shapes.Count; rightIndex++) {
                    PowerPointShape right = shapes[rightIndex];
                    if (ShouldSkipCollisionShape(right)) continue;

                    PowerPointLayoutBox leftBounds = left.Bounds;
                    PowerPointLayoutBox rightBounds = right.Bounds;
                    if (options.IgnoreContainedShapeCollisions &&
                        (Contains(leftBounds, rightBounds, tolerance) || Contains(rightBounds, leftBounds, tolerance))) {
                        continue;
                    }

                    long intersectionWidth = Math.Min(leftBounds.Right.ToEmus(), rightBounds.Right.ToEmus()) -
                                             Math.Max(leftBounds.Left.ToEmus(), rightBounds.Left.ToEmus());
                    long intersectionHeight = Math.Min(leftBounds.Bottom.ToEmus(), rightBounds.Bottom.ToEmus()) -
                                              Math.Max(leftBounds.Top.ToEmus(), rightBounds.Top.ToEmus());
                    if (intersectionWidth <= tolerance || intersectionHeight <= tolerance) {
                        continue;
                    }

                    double intersectionArea = intersectionWidth * (double)intersectionHeight;
                    double smallerArea = Math.Min(leftBounds.Width.ToEmus() * (double)leftBounds.Height.ToEmus(),
                        rightBounds.Width.ToEmus() * (double)rightBounds.Height.ToEmus());
                    if (smallerArea <= 0D || intersectionArea / smallerArea < options.MinimumCollisionOverlapRatio) {
                        continue;
                    }

                    string rightLabel = string.IsNullOrWhiteSpace(right.Name)
                        ? "shape " + rightIndex.ToString(CultureInfo.InvariantCulture)
                        : "'" + right.Name + "'";
                    findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Warning, "layout-shape-collision",
                        "Shape significantly overlaps " + rightLabel + ".", slideIndex,
                        containingShapeIndex ?? leftIndex, left,
                        leftBounds));
                }
            }
        }

        private static bool ShouldSkipCollisionShape(PowerPointShape shape) =>
            shape.Hidden || shape.ShapeContentType == PowerPointShapeContentType.Connector ||
            shape.ShapeContentType == PowerPointShapeContentType.Group;

        private static bool Contains(PowerPointLayoutBox outer, PowerPointLayoutBox inner, double tolerance) =>
            inner.Left.ToEmus() >= outer.Left.ToEmus() - tolerance && inner.Top.ToEmus() >= outer.Top.ToEmus() - tolerance &&
            inner.Right.ToEmus() <= outer.Right.ToEmus() + tolerance && inner.Bottom.ToEmus() <= outer.Bottom.ToEmus() + tolerance;

        private static void InspectVisualSnapshot(PowerPointSlide slide, int slideIndex,
            IList<PowerPointDeckPreflightFinding> findings) {
            try {
                PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
                bool spacingReported = false;
                for (int index = 0; index < snapshot.Diagnostics.Count; index++) {
                    OfficeImageExportDiagnostic diagnostic = snapshot.Diagnostics[index];
                    PowerPointDeckPreflightSeverity severity = diagnostic.Severity switch {
                        OfficeImageExportDiagnosticSeverity.Error => PowerPointDeckPreflightSeverity.Error,
                        OfficeImageExportDiagnosticSeverity.Warning => PowerPointDeckPreflightSeverity.Warning,
                        _ => PowerPointDeckPreflightSeverity.Info
                    };
                    // Tracking is kept in the file; only the preview approximates it. That says nothing about the
                    // deck, so it is reported once per slide as information rather than once per text frame.
                    if (diagnostic.Code == "powerpoint-character-spacing-approximated") {
                        if (spacingReported) continue;
                        spacingReported = true;
                        severity = PowerPointDeckPreflightSeverity.Info;
                    }
                    findings.Add(new PowerPointDeckPreflightFinding(severity,
                        "VisualSnapshot." + diagnostic.Code, diagnostic.Message, slideIndex,
                        shapeName: diagnostic.Source));
                }
            } catch (Exception exception) when (exception is ArgumentException ||
                                                exception is InvalidOperationException ||
                                                exception is NotSupportedException) {
                findings.Add(new PowerPointDeckPreflightFinding(PowerPointDeckPreflightSeverity.Error,
                    "visual-snapshot-failed", "Shared visual snapshot failed: " + exception.Message, slideIndex));
            }
        }

        private static PowerPointDeckPreflightFinding CreateFinding(PowerPointDeckPreflightSeverity severity,
            string code, string message, int slideIndex, int shapeIndex, PowerPointShape shape,
            PowerPointLayoutBox bounds, double? resolvedFontSize = null) =>
            new PowerPointDeckPreflightFinding(severity, code, message, slideIndex, shapeIndex, shape.Id,
                shape.Name, bounds, resolvedFontSize);
    }
}
