using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Inspects authored slide bounds, measured text fit, peer collisions, image relationships, and
        ///     shared visual-snapshot diagnostics without changing the presentation.
        /// </summary>
        public PowerPointDeckPreflightReport InspectPreflight(PowerPointDeckPreflightOptions? options = null) {
            ThrowIfDisposed();
            PowerPointDeckPreflightOptions resolved = options?.Clone() ?? new PowerPointDeckPreflightOptions();
            var findings = new List<PowerPointDeckPreflightFinding>();
            long slideWidth = SlideSize.WidthEmus;
            long slideHeight = SlideSize.HeightEmus;
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
                new PowerPointLayoutBox(0L, 0L, slideWidth, slideHeight), options, findings, null,
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
                    IReadOnlyList<PowerPointShape> children = slide.GetGroupChildren(groupShape);
                    InspectShapeTree(slide, children, slideIndex, slide.GetGroupChildBounds(groupShape),
                        options, findings, reportShapeIndex, groupDepth + 1, ref inspectedShapeCount);
                }
            }

            if (options.DetectShapeCollisions) {
                InspectCollisions(shapes, slideIndex, options, findings, containingShapeIndex);
            }
        }

        private static void InspectBounds(PowerPointShape shape, int shapeIndex, int slideIndex,
            PowerPointLayoutBox canvas, PowerPointDeckPreflightOptions options,
            IList<PowerPointDeckPreflightFinding> findings) {
            PowerPointLayoutBox bounds = shape.Bounds;
            if (bounds.Width <= 0L || bounds.Height <= 0L) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "Layout.InvalidBounds",
                    "Shape has a non-positive width or height.", slideIndex, shapeIndex, shape, bounds));
                return;
            }

            if (bounds.Left < canvas.Left || bounds.Top < canvas.Top ||
                bounds.Right > canvas.Right || bounds.Bottom > canvas.Bottom) {
                long bleed = PowerPointUnits.FromPoints(options.MaximumDecorativeBleedPoints);
                bool allowedDecorativeBleed = options.AllowDecorativeShapeBleed &&
                    shape is PowerPointAutoShape &&
                    bounds.Left >= canvas.Left - bleed && bounds.Top >= canvas.Top - bleed &&
                    bounds.Right <= canvas.Right + bleed && bounds.Bottom <= canvas.Bottom + bleed &&
                    bounds.Right > canvas.Left && bounds.Bottom > canvas.Top &&
                    bounds.Left < canvas.Right && bounds.Top < canvas.Bottom;
                if (allowedDecorativeBleed) {
                    return;
                }
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "Layout.ShapeOffSlide",
                    "Shape extends beyond its slide or group canvas.", slideIndex, shapeIndex, shape, bounds));
            }
        }

        private static void InspectText(PowerPointTextBox textBox, int shapeIndex, int slideIndex,
            PowerPointDeckPreflightOptions options, IList<PowerPointDeckPreflightFinding> findings) {
            string text = textBox.Text;
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            double authoredFontSize = ResolveFontSize(textBox, options.DefaultFontSizePoints);
            string fontName = ResolveFontName(textBox);
            OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
            if (textBox.Bold) fontStyle |= OfficeFontStyle.Bold;
            if (textBox.Italic) fontStyle |= OfficeFontStyle.Italic;
            var font = new OfficeFontInfo(fontName, authoredFontSize, fontStyle);
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(font);
            Func<string?, double, double> measure = (value, size) => {
                OfficeTextMeasurementStyle style = measurer.CreateStyle(font.WithSize(size));
                return measurer.MeasureWidth(value, style) * 72D / OfficeTextMeasurer.DefaultDpi;
            };

            double contentWidth = Math.Max(1D, textBox.WidthPoints -
                (textBox.TextMarginLeftPoints ?? 7.2D) - (textBox.TextMarginRightPoints ?? 7.2D));
            double contentHeight = Math.Max(1D, textBox.HeightPoints -
                (textBox.TextMarginTopPoints ?? 3.6D) - (textBox.TextMarginBottomPoints ?? 3.6D));
            double resolvedFontSize = authoredFontSize;

            if (textBox.TextAutoFit == PowerPointTextAutoFit.Normal) {
                OfficeTextBlockLayout fitted = OfficeTextLayoutEngine.FitWrappedText(text, authoredFontSize,
                    contentWidth, contentHeight, 1.2D, 1D, measure);
                resolvedFontSize = fitted.FontSize;
                double? authoredScale = textBox.TextAutoFitOptions?.FontScalePercent;
                if (authoredScale.HasValue) {
                    resolvedFontSize = Math.Min(resolvedFontSize, authoredFontSize * authoredScale.Value / 100D);
                }
            }

            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(text, resolvedFontSize,
                contentWidth, contentHeight, 1.2D, 1D, measure, wrap: true, forceSingleLine: false,
                shrinkToFit: false);

            if (options.DetectTextOverflow && layout.Clipped &&
                textBox.TextAutoFit != PowerPointTextAutoFit.Shape) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "Text.Clipped",
                    "Measured text does not fit inside the authored text box.", slideIndex, shapeIndex, textBox,
                    textBox.Bounds, resolvedFontSize));
            }

            if (options.DetectUnreadableFontReduction && resolvedFontSize < options.MinimumReadableFontSizePoints) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Warning, "Text.UnreadableFontReduction",
                    "Text fit resolves to " + resolvedFontSize.ToString("0.##", CultureInfo.InvariantCulture) +
                    " pt, below the configured " +
                    options.MinimumReadableFontSizePoints.ToString("0.##", CultureInfo.InvariantCulture) +
                    " pt threshold.", slideIndex, shapeIndex, textBox, textBox.Bounds, resolvedFontSize));
            }
        }

        private static double ResolveFontSize(PowerPointTextBox textBox, double fallback) {
            double maximum = 0D;
            for (int paragraphIndex = 0; paragraphIndex < textBox.Paragraphs.Count; paragraphIndex++) {
                IReadOnlyList<PowerPointTextRun> runs = textBox.Paragraphs[paragraphIndex].Runs;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    maximum = Math.Max(maximum, runs[runIndex].FontSize ?? 0D);
                }
            }

            return maximum > 0D ? maximum : textBox.FontSize ?? fallback;
        }

        private static string ResolveFontName(PowerPointTextBox textBox) {
            for (int paragraphIndex = 0; paragraphIndex < textBox.Paragraphs.Count; paragraphIndex++) {
                IReadOnlyList<PowerPointTextRun> runs = textBox.Paragraphs[paragraphIndex].Runs;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    if (!string.IsNullOrWhiteSpace(runs[runIndex].FontName)) {
                        return runs[runIndex].FontName!;
                    }
                }
            }

            return string.IsNullOrWhiteSpace(textBox.FontName) ? "Aptos" : textBox.FontName!;
        }

        private static void InspectPicture(PowerPointPicture picture, int shapeIndex, int slideIndex,
            IList<PowerPointDeckPreflightFinding> findings) {
            try {
                if (string.IsNullOrWhiteSpace(picture.ContentType)) {
                    findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "Asset.MissingImagePart",
                        "Picture does not resolve to an embedded image part.", slideIndex, shapeIndex, picture,
                        picture.Bounds));
                }
            } catch (Exception exception) when (exception is ArgumentOutOfRangeException ||
                                                exception is KeyNotFoundException ||
                                                exception is InvalidOperationException) {
                findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Error, "Asset.BrokenImageRelationship",
                    "Picture relationship cannot be resolved: " + exception.Message, slideIndex, shapeIndex,
                    picture, picture.Bounds));
            }
        }

        private static void InspectCollisions(IReadOnlyList<PowerPointShape> shapes, int slideIndex,
            PowerPointDeckPreflightOptions options, IList<PowerPointDeckPreflightFinding> findings,
            int? containingShapeIndex) {
            double tolerance = PowerPointUnits.FromPoints(options.CollisionTolerancePoints);
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

                    long intersectionWidth = Math.Min(leftBounds.Right, rightBounds.Right) -
                                             Math.Max(leftBounds.Left, rightBounds.Left);
                    long intersectionHeight = Math.Min(leftBounds.Bottom, rightBounds.Bottom) -
                                              Math.Max(leftBounds.Top, rightBounds.Top);
                    if (intersectionWidth <= tolerance || intersectionHeight <= tolerance) {
                        continue;
                    }

                    double intersectionArea = intersectionWidth * (double)intersectionHeight;
                    double smallerArea = Math.Min(leftBounds.Width * (double)leftBounds.Height,
                        rightBounds.Width * (double)rightBounds.Height);
                    if (smallerArea <= 0D || intersectionArea / smallerArea < options.MinimumCollisionOverlapRatio) {
                        continue;
                    }

                    string rightLabel = string.IsNullOrWhiteSpace(right.Name)
                        ? "shape " + rightIndex.ToString(CultureInfo.InvariantCulture)
                        : "'" + right.Name + "'";
                    findings.Add(CreateFinding(PowerPointDeckPreflightSeverity.Warning, "Layout.ShapeCollision",
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
            inner.Left >= outer.Left - tolerance && inner.Top >= outer.Top - tolerance &&
            inner.Right <= outer.Right + tolerance && inner.Bottom <= outer.Bottom + tolerance;

        private static void InspectVisualSnapshot(PowerPointSlide slide, int slideIndex,
            IList<PowerPointDeckPreflightFinding> findings) {
            try {
                PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
                for (int index = 0; index < snapshot.Diagnostics.Count; index++) {
                    OfficeImageExportDiagnostic diagnostic = snapshot.Diagnostics[index];
                    PowerPointDeckPreflightSeverity severity = diagnostic.Severity switch {
                        OfficeImageExportDiagnosticSeverity.Error => PowerPointDeckPreflightSeverity.Error,
                        OfficeImageExportDiagnosticSeverity.Warning => PowerPointDeckPreflightSeverity.Warning,
                        _ => PowerPointDeckPreflightSeverity.Info
                    };
                    findings.Add(new PowerPointDeckPreflightFinding(severity,
                        "VisualSnapshot." + diagnostic.Code, diagnostic.Message, slideIndex,
                        shapeName: diagnostic.Source));
                }
            } catch (Exception exception) when (exception is ArgumentException ||
                                                exception is InvalidOperationException ||
                                                exception is NotSupportedException) {
                findings.Add(new PowerPointDeckPreflightFinding(PowerPointDeckPreflightSeverity.Error,
                    "VisualSnapshot.Failed", "Shared visual snapshot failed: " + exception.Message, slideIndex));
            }
        }

        private static PowerPointDeckPreflightFinding CreateFinding(PowerPointDeckPreflightSeverity severity,
            string code, string message, int slideIndex, int shapeIndex, PowerPointShape shape,
            PowerPointLayoutBox bounds, double? resolvedFontSize = null) =>
            new PowerPointDeckPreflightFinding(severity, code, message, slideIndex, shapeIndex, shape.Id,
                shape.Name, bounds, resolvedFontSize);
    }
}
