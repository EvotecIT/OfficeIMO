using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfStaticFormRecognizer {
    private static List<PdfPageVisualPrimitive> ExpandIndependentPathStrokes(
        IReadOnlyList<PdfPageVisualPrimitive> primitives, ref long candidateScanWork,
        int maxCandidateScanWork, CancellationToken cancellationToken) {
        var expanded = new List<PdfPageVisualPrimitive>(primitives.Count);
        foreach (PdfPageVisualPrimitive primitive in primitives) {
            cancellationToken.ThrowIfCancellationRequested();
            if (primitive.Kind != PdfPageVisualPrimitiveKind.Path || !primitive.HasStrokePaint ||
                primitive.PathCommands.Count == 0) {
                expanded.Add(primitive);
                continue;
            }
            candidateScanWork = checked(candidateScanWork + primitive.PathCommands.Count);
            if (candidateScanWork > maxCandidateScanWork) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                    maxCandidateScanWork, candidateScanWork);
            }
            var parts = new List<PdfPageVisualPrimitive>();
            var subpath = new List<OfficePathCommand>();
            bool supported = true;
            foreach (OfficePathCommand command in primitive.PathCommands) {
                if (command.Kind == OfficePathCommandKind.MoveTo && subpath.Count > 0) {
                    supported &= TryAddSimpleSubpath(primitive, subpath, parts);
                    subpath.Clear();
                }
                subpath.Add(command);
            }
            if (subpath.Count > 0) supported &= TryAddSimpleSubpath(primitive, subpath, parts);
            if (supported && parts.Count > 0) expanded.AddRange(parts);
            else expanded.Add(primitive);
        }
        return expanded;
    }

    private static bool TryAddSimpleSubpath(PdfPageVisualPrimitive source,
        List<OfficePathCommand> commands, List<PdfPageVisualPrimitive> parts) {
        if (commands[0].Kind != OfficePathCommandKind.MoveTo) return false;
        PdfPageVisualPrimitive part;
        if (commands.Count == 2 && commands[1].Kind == OfficePathCommandKind.LineTo &&
            Math.Abs(commands[0].Point.Y - commands[1].Point.Y) <= 1D) {
            part = PdfPageVisualPrimitive.Line(commands[0].Point.X, commands[0].Point.Y,
                commands[1].Point.X, commands[1].Point.Y, source.StrokeColor,
                source.StrokeGradient, source.StrokeRadialGradient, source.StrokeWidth,
                source.StrokeDashStyle, source.StrokeLineCap, source.StrokeLineJoin,
                source.StrokeOpacity, source.ClipPath, source.PaintOrder,
                source.StrokeTilingPattern, source.StrokeDashPattern);
        } else if (commands.Count == 5 &&
                   commands[1].Kind == OfficePathCommandKind.LineTo &&
                   commands[2].Kind == OfficePathCommandKind.LineTo &&
                   commands[3].Kind == OfficePathCommandKind.LineTo &&
                   commands[4].Kind == OfficePathCommandKind.Close &&
                   IsAxisAlignedRectangle(commands)) {
            double left = Math.Min(commands[0].Point.X, commands[2].Point.X);
            double top = Math.Min(commands[0].Point.Y, commands[2].Point.Y);
            double right = Math.Max(commands[0].Point.X, commands[2].Point.X);
            double bottom = Math.Max(commands[0].Point.Y, commands[2].Point.Y);
            part = PdfPageVisualPrimitive.Rectangle(left, top, right - left, bottom - top,
                source.FillColor, source.FillGradient, source.FillRadialGradient,
                source.StrokeColor, source.StrokeGradient, source.StrokeRadialGradient,
                source.StrokeWidth, source.StrokeDashStyle, source.StrokeLineCap,
                source.StrokeLineJoin, source.FillOpacity, source.StrokeOpacity,
                source.ClipPath, source.PaintOrder, source.FillTilingPattern,
                source.StrokeTilingPattern, source.StrokeDashPattern);
        } else return false;
        if (source.ContentOrderKey is PdfContentOrderKey key) part = part.WithContentOrderKey(key);
        parts.Add(part.WithSourceOperatorIndex(source.SourceOperatorIndex));
        return true;
    }

    private static bool IsAxisAlignedRectangle(List<OfficePathCommand> commands) {
        OfficePoint p0 = commands[0].Point;
        OfficePoint p1 = commands[1].Point;
        OfficePoint p2 = commands[2].Point;
        OfficePoint p3 = commands[3].Point;
        const double tolerance = 0.000001D;
        return (Math.Abs(p0.X - p1.X) <= tolerance && Math.Abs(p1.Y - p2.Y) <= tolerance &&
                Math.Abs(p2.X - p3.X) <= tolerance && Math.Abs(p3.Y - p0.Y) <= tolerance ||
                Math.Abs(p0.Y - p1.Y) <= tolerance && Math.Abs(p1.X - p2.X) <= tolerance &&
                Math.Abs(p2.Y - p3.Y) <= tolerance && Math.Abs(p3.X - p0.X) <= tolerance) &&
            Math.Abs(p0.X - p2.X) > tolerance && Math.Abs(p0.Y - p2.Y) > tolerance;
    }

}
