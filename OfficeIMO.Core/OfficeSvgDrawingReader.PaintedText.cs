using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryAddPaintedTextRun(
        OfficeDrawing drawing,
        SvgTextRun run,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        if (run.FontProgram is not IOfficeBoundedFontProgram bounded) return false;
        int remainingCommands = MaximumSvgPathCommands - pathCommands;
        if (remainingCommands < 8) {
            pathCommandLimitExceeded = true;
            return false;
        }

        double lineHeight = bounded.LineHeight(run.FontSize);
        double baselineOffset = bounded is IOfficeFontBaselineMetrics metrics
            ? metrics.BaselineOffset(run.FontSize)
            : lineHeight * 0.8D;
        if (double.IsNaN(lineHeight) || double.IsInfinity(lineHeight) || lineHeight <= 0D
            || double.IsNaN(baselineOffset) || double.IsInfinity(baselineOffset)) return false;
        baselineOffset = Math.Max(0D, Math.Min(lineHeight, baselineOffset));
        int pointAllowance = Math.Max(1, remainingCommands / 2);
        List<List<OfficePoint>> contours;
        try {
            contours = bounded is IOfficeCffBoundedFontProgram cff
                ? cff.GetTextContoursBounded(
                    run.Text,
                    run.X,
                    run.Baseline - baselineOffset,
                    run.FontSize,
                    pointAllowance,
                    CancellationToken.None,
                    new OfficeCffOperationBudget())
                : bounded.GetTextContoursBounded(
                    run.Text,
                    run.X,
                    run.Baseline - baselineOffset,
                    run.FontSize,
                    pointAllowance,
                    CancellationToken.None);
        } catch (InvalidOperationException) {
            pathCommandLimitExceeded = true;
            return false;
        } catch (ArgumentException) {
            return false;
        }
        if (contours.Count == 0) return true;

        double minimumX = double.PositiveInfinity;
        double minimumY = double.PositiveInfinity;
        double maximumX = double.NegativeInfinity;
        double maximumY = double.NegativeInfinity;
        var transformedContours = new List<List<OfficePoint>>(contours.Count);
        foreach (List<OfficePoint> contour in contours) {
            if (contour.Count < 3) continue;
            var transformed = new List<OfficePoint>(contour.Count);
            foreach (OfficePoint point in contour) {
                var scaled = new OfficePoint(
                    run.X + ((point.X - run.X) * run.GlyphScale),
                    point.Y);
                transformed.Add(scaled);
                minimumX = Math.Min(minimumX, scaled.X);
                minimumY = Math.Min(minimumY, scaled.Y);
                maximumX = Math.Max(maximumX, scaled.X);
                maximumY = Math.Max(maximumY, scaled.Y);
            }
            transformedContours.Add(transformed);
        }
        if (transformedContours.Count == 0) return true;
        if (minimumX < 0D || minimumY < 0D || maximumX > drawing.Width || maximumY > drawing.Height) return false;

        var commands = new List<OfficePathCommand>();
        foreach (List<OfficePoint> contour in transformedContours) {
            commands.Add(OfficePathCommand.MoveTo(contour[0].X - minimumX, contour[0].Y - minimumY));
            for (int index = 1; index < contour.Count; index++) {
                commands.Add(OfficePathCommand.LineTo(contour[index].X - minimumX, contour[index].Y - minimumY));
            }
            commands.Add(OfficePathCommand.Close());
        }
        if (commands.Count > remainingCommands) {
            pathCommandLimitExceeded = true;
            return false;
        }
        pathCommands += commands.Count;

        OfficeShape outline = OfficeShape.Path(
            Math.Max(0.0001D, maximumX - minimumX),
            Math.Max(0.0001D, maximumY - minimumY),
            commands);
        ApplyPaint(outline, run.Style);
        var positioned = new OfficeDrawingShape(outline, minimumX, minimumY);
        ApplyDeferredPaint(
            positioned.Shape,
            run.Style,
            positioned.X,
            positioned.Y,
            drawing.Width,
            drawing.Height,
            viewX,
            viewY,
            ref unsupported);

        OfficeTransform textTransform = Math.Abs(run.RotationDegrees) <= 0.0000001D
            ? run.Transform
            : OfficeTransform.RotateDegrees(run.RotationDegrees, run.RotationCenterX, run.RotationCenterY).Then(run.Transform);
        ApplyTransform(positioned, textTransform);
        if (run.Style.StrokePattern != null) {
            unsupported++;
            positioned.Shape.StrokeColor = null;
            positioned.Shape.StrokeGradient = null;
            positioned.Shape.StrokeRadialGradient = null;
        }

        bool hasPattern = TryAddSvgPatternFill(
            run.Style.FillPattern,
            positioned,
            drawing,
            run.Style,
            paintServers,
            references,
            textTransform,
            viewX,
            viewY,
            maximumElements,
            maximumViewportDimension,
            maximumViewportPixels,
            depth,
            ref visited,
            ref pathCommands,
            ref pathCommandLimitExceeded,
            ref unsupported,
            out OfficeDrawing? patternLayer);

        var paint = new OfficeDrawing(drawing.Width, drawing.Height);
        paint.Fonts.AddRange(drawing.Fonts);
        if (hasPattern && patternLayer != null) paint.AddEffectDrawing(patternLayer, OfficeTransform.Identity);
        paint.AddShape(positioned.Shape, positioned.X, positioned.Y);
        OfficePoint anchor = textTransform.TransformPoint(new OfficePoint(run.X, run.Baseline));
        drawing.AddActualTextDrawing(run.Text, paint, anchor.X, anchor.Y);
        return true;
    }
}
