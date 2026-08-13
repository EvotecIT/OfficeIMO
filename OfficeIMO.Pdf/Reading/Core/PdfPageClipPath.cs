using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly partial struct PdfPageClipPath {
    private PdfPageClipPath(double x, double y, double width, double height, bool isRectangle, OfficeFillRule fillRule, IReadOnlyList<OfficePathCommand> commands, bool isExact = true) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        IsRectangle = isRectangle;
        FillRule = fillRule;
        Commands = commands;
        IsExact = isExact;
    }

    public static PdfPageClipPath Rectangle(double x, double y, double width, double height) =>
        new PdfPageClipPath(x, y, width, height, true, OfficeFillRule.EvenOdd, Array.Empty<OfficePathCommand>());

    public static PdfPageClipPath ResolveActiveClip(PdfPageClipPath? activeClipPath, PdfPageClipPath clipPath) {
        if (!activeClipPath.HasValue) {
            return clipPath;
        }

        PdfPageClipPath active = activeClipPath.Value;
        if (!active.IsRectangle || !clipPath.IsRectangle) {
            if (active.IsRectangle) {
                PdfPageClipPath resolved = IntersectClipBounds(active, clipPath, out PdfPageClipPath intersection)
                    ? IntersectPathWithRectangle(clipPath, active, intersection)
                    : Rectangle(Math.Max(active.X, clipPath.X), Math.Max(active.Y, clipPath.Y), 0D, 0D);
                return resolved.WithExactness(resolved.IsExact && active.IsExact && clipPath.IsExact);
            }

            if (clipPath.IsRectangle) {
                PdfPageClipPath resolved = IntersectClipBounds(active, clipPath, out PdfPageClipPath intersection)
                    ? IntersectPathWithRectangle(active, clipPath, intersection)
                    : Rectangle(Math.Max(active.X, clipPath.X), Math.Max(active.Y, clipPath.Y), 0D, 0D);
                return resolved.WithExactness(resolved.IsExact && active.IsExact && clipPath.IsExact);
            }

            if (!IntersectClipBounds(active, clipPath, out PdfPageClipPath pathIntersection)) {
                return Rectangle(Math.Max(active.X, clipPath.X), Math.Max(active.Y, clipPath.Y), 0D, 0D)
                    .WithExactness(active.IsExact && clipPath.IsExact);
            }

            PdfPageClipPath pathResult = CanServeAsExactPathClip(clipPath) || !CanServeAsExactPathClip(active)
                ? IntersectPathWithPath(active, clipPath, pathIntersection)
                : IntersectPathWithPath(clipPath, active, pathIntersection);
            return pathResult.WithExactness(pathResult.IsExact && active.IsExact && clipPath.IsExact);
        }

        PdfPageClipPath rectangleResult = IntersectClipBounds(active, clipPath, out PdfPageClipPath rectangleIntersection)
            ? rectangleIntersection
            : Rectangle(Math.Max(active.X, clipPath.X), Math.Max(active.Y, clipPath.Y), 0D, 0D);
        return rectangleResult.WithExactness(active.IsExact && clipPath.IsExact);
    }

    public static bool TryCombineTextClippingPaths(IReadOnlyList<PdfPageClipPath> paths, out PdfPageClipPath clipPath) {
        clipPath = default;
        if (paths.Count == 0) return false;
        if (paths.Count == 1) {
            clipPath = paths[0];
            return true;
        }

        var commands = new List<OfficePathCommand>();
        for (int i = 0; i < paths.Count; i++) {
            PdfPageClipPath path = paths[i];
            if (!path.IsRectangle) {
                commands.AddRange(path.Commands);
                continue;
            }

            double right = path.X + path.Width;
            double bottom = path.Y + path.Height;
            commands.Add(OfficePathCommand.MoveTo(new OfficePoint(path.X, path.Y)));
            commands.Add(OfficePathCommand.LineTo(new OfficePoint(right, path.Y)));
            commands.Add(OfficePathCommand.LineTo(new OfficePoint(right, bottom)));
            commands.Add(OfficePathCommand.LineTo(new OfficePoint(path.X, bottom)));
            commands.Add(OfficePathCommand.Close());
        }

        return TryCreatePath(commands, OfficeFillRule.NonZero, out clipPath);
    }

    private static bool IntersectClipBounds(PdfPageClipPath first, PdfPageClipPath second, out PdfPageClipPath intersection) {
        double left = Math.Max(first.X, second.X);
        double top = Math.Max(first.Y, second.Y);
        double right = Math.Min(first.X + first.Width, second.X + second.Width);
        double bottom = Math.Min(first.Y + first.Height, second.Y + second.Height);
        double width = right - left;
        double height = bottom - top;
        if (width <= 0D || height <= 0D) {
            intersection = default;
            return false;
        }

        intersection = Rectangle(left, top, width, height);
        return true;
    }

    private static PdfPageClipPath IntersectPathWithRectangle(PdfPageClipPath pathClip, PdfPageClipPath rectangleClip, PdfPageClipPath intersection) {
        List<OfficePathCommand> clippedCommands = ClipPathCommandsToRectangle(pathClip.Commands, rectangleClip);
        PdfPageClipPath result = clippedCommands.Count > 0 && TryCreatePath(clippedCommands, pathClip.FillRule, out PdfPageClipPath clippedPath)
            ? clippedPath
            : Rectangle(intersection.X, intersection.Y, 0D, 0D);
        return result.WithExactness(pathClip.IsExact &&
            !ContainsCurve(pathClip.Commands) &&
            HasRepresentableClippedContours(pathClip, result));
    }

    private static PdfPageClipPath IntersectPathWithPath(PdfPageClipPath active, PdfPageClipPath next, PdfPageClipPath intersection) {
        bool isExact = active.IsExact && next.IsExact &&
            !ContainsCurve(active.Commands) &&
            !ContainsCurve(next.Commands);
        List<List<OfficePoint>> subjectContours = FlattenPathContours(active.Commands);
        List<List<OfficePoint>> clipContours = FlattenPathContours(next.Commands);
        if (subjectContours.Count == 0 || clipContours.Count == 0) {
            return Rectangle(intersection.X, intersection.Y, 0D, 0D).WithExactness(isExact);
        }

        var intersectedContours = new List<List<OfficePoint>>();
        bool canClipPerContour = clipContours.All(IsConvexContour) && !HasOverlappingContourBounds(clipContours);
        if (!canClipPerContour) {
            // Exact arbitrary path intersection needs a general polygon boolean engine.
            // Preserve a conservative superset so unsupported clip complexity cannot
            // suppress visible-content reporting or discard the rendered element.
            return intersection.WithExactness(false);
        }

        for (int i = 0; i < subjectContours.Count; i++) {
            for (int clipIndex = 0; clipIndex < clipContours.Count; clipIndex++) {
                List<OfficePoint> clipped = ClipPolygonToConvexPolygon(subjectContours[i], clipContours[clipIndex]);
                if (clipped.Count >= 3) {
                    intersectedContours.Add(clipped);
                }
            }
        }

        List<OfficePathCommand> commands = BuildClosedContourCommands(intersectedContours);
        PdfPageClipPath result = commands.Count > 0 && TryCreatePath(commands, active.FillRule, out PdfPageClipPath path)
            ? path
            : Rectangle(intersection.X, intersection.Y, 0D, 0D);
        return result.WithExactness(isExact);
    }

    private static bool CanServeAsExactPathClip(PdfPageClipPath path) {
        if (path.IsRectangle) return true;
        if (!path.IsExact || ContainsCurve(path.Commands)) return false;
        List<List<OfficePoint>> contours = FlattenPathContours(path.Commands);
        return contours.Count > 0 && contours.All(IsConvexContour) && !HasOverlappingContourBounds(contours);
    }

    private static bool ContainsCurve(IReadOnlyList<OfficePathCommand> commands) {
        for (int i = 0; i < commands.Count; i++) {
            if (commands[i].Kind == OfficePathCommandKind.QuadraticBezierTo ||
                commands[i].Kind == OfficePathCommandKind.CubicBezierTo) return true;
        }
        return false;
    }

    private static bool HasRepresentableClippedContours(PdfPageClipPath source, PdfPageClipPath clipped) {
        if ((clipped.IsRectangle && clipped.Width <= 0D) || clipped.Height <= 0D) return true;
        List<List<OfficePoint>> sourceContours = FlattenPathContours(source.Commands);
        List<List<OfficePoint>> clippedContours = clipped.IsRectangle
            ? new List<List<OfficePoint>> {
                new List<OfficePoint> {
                    new OfficePoint(clipped.X, clipped.Y),
                    new OfficePoint(clipped.X + clipped.Width, clipped.Y),
                    new OfficePoint(clipped.X + clipped.Width, clipped.Y + clipped.Height),
                    new OfficePoint(clipped.X, clipped.Y + clipped.Height)
                }
            }
            : FlattenPathContours(clipped.Commands);
        if (sourceContours.Count == 0 || clippedContours.Count == 0) return false;
        for (int contourIndex = 0; contourIndex < clippedContours.Count; contourIndex++) {
            List<OfficePoint> contour = clippedContours[contourIndex];
            for (int pointIndex = 0; pointIndex < contour.Count; pointIndex++) {
                OfficePoint start = contour[pointIndex];
                OfficePoint end = contour[(pointIndex + 1) % contour.Count];
                if (!IsSegmentWithinFilledArea(sourceContours, source.FillRule, start, end)) return false;
            }
        }
        return true;
    }

    private static bool ContainsFilledPoint(List<List<OfficePoint>> contours, OfficeFillRule fillRule, OfficePoint point) {
        int winding = 0;
        bool inside = false;
        for (int contourIndex = 0; contourIndex < contours.Count; contourIndex++) {
            List<OfficePoint> contour = contours[contourIndex];
            for (int pointIndex = 0; pointIndex < contour.Count; pointIndex++) {
                OfficePoint start = contour[pointIndex];
                OfficePoint end = contour[(pointIndex + 1) % contour.Count];
                if (IsPointOnSegment(point, start, end)) return true;
                bool crosses = start.Y <= point.Y ? end.Y > point.Y : end.Y <= point.Y;
                if (!crosses) continue;
                double cross = ((end.X - start.X) * (point.Y - start.Y)) - ((point.X - start.X) * (end.Y - start.Y));
                if (fillRule == OfficeFillRule.EvenOdd) {
                    if ((end.Y > start.Y && cross > 0D) || (end.Y < start.Y && cross < 0D)) inside = !inside;
                } else if (end.Y > start.Y && cross > 0D) {
                    winding++;
                } else if (end.Y < start.Y && cross < 0D) {
                    winding--;
                }
            }
        }
        return fillRule == OfficeFillRule.EvenOdd ? inside : winding != 0;
    }

    private static bool IsPointOnSegment(OfficePoint point, OfficePoint start, OfficePoint end) {
        double cross = ((end.X - start.X) * (point.Y - start.Y)) - ((end.Y - start.Y) * (point.X - start.X));
        if (Math.Abs(cross) > 0.001D) return false;
        return point.X >= Math.Min(start.X, end.X) - 0.001D &&
            point.X <= Math.Max(start.X, end.X) + 0.001D &&
            point.Y >= Math.Min(start.Y, end.Y) - 0.001D &&
            point.Y <= Math.Max(start.Y, end.Y) + 0.001D;
    }

    private static bool IsConvexContour(List<OfficePoint> contour) {
        if (contour.Count < 3) {
            return false;
        }

        double sign = 0D;
        for (int i = 0; i < contour.Count; i++) {
            OfficePoint a = contour[i];
            OfficePoint b = contour[(i + 1) % contour.Count];
            OfficePoint c = contour[(i + 2) % contour.Count];
            double cross = ((b.X - a.X) * (c.Y - b.Y)) - ((b.Y - a.Y) * (c.X - b.X));
            if (cross == 0D) {
                continue;
            }

            double currentSign = Math.Sign(cross);
            if (sign == 0D) {
                sign = currentSign;
            } else if (Math.Sign(sign) != Math.Sign(currentSign)) {
                return false;
            }
        }

        return sign != 0D;
    }

    private static bool HasOverlappingContourBounds(List<List<OfficePoint>> contours) {
        for (int i = 0; i < contours.Count; i++) {
            GetContourBounds(contours[i], out double left, out double top, out double right, out double bottom);
            for (int j = i + 1; j < contours.Count; j++) {
                GetContourBounds(contours[j], out double otherLeft, out double otherTop, out double otherRight, out double otherBottom);
                if (left < otherRight && right > otherLeft && top < otherBottom && bottom > otherTop) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void GetContourBounds(List<OfficePoint> contour, out double left, out double top, out double right, out double bottom) {
        left = right = contour[0].X;
        top = bottom = contour[0].Y;
        for (int i = 1; i < contour.Count; i++) {
            OfficePoint point = contour[i];
            left = Math.Min(left, point.X);
            top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X);
            bottom = Math.Max(bottom, point.Y);
        }
    }

    private static List<List<OfficePoint>> FlattenPathContours(IReadOnlyList<OfficePathCommand> commands) {
        var contours = new List<List<OfficePoint>>();
        List<OfficePoint>? current = null;
        OfficePoint currentPoint = default;
        bool hasCurrentPoint = false;
        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    AddFlattenedContour(contours, current);
                    currentPoint = command.Point;
                    current = new List<OfficePoint> { currentPoint };
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.LineTo:
                    EnsureContour(ref current, currentPoint, hasCurrentPoint);
                    currentPoint = command.Point;
                    current!.Add(currentPoint);
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    EnsureContour(ref current, currentPoint, hasCurrentPoint);
                    current!.AddRange(OfficeGeometry.CreateQuadraticBezierPoints(currentPoint, command.ControlPoint1, command.Point, 24));
                    currentPoint = command.Point;
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    EnsureContour(ref current, currentPoint, hasCurrentPoint);
                    current!.AddRange(OfficeGeometry.CreateCubicBezierPoints(currentPoint, command.ControlPoint1, command.ControlPoint2, command.Point, 24));
                    currentPoint = command.Point;
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.Close:
                    if (current != null && current.Count > 0) {
                        currentPoint = current[0];
                        hasCurrentPoint = true;
                    }
                    AddFlattenedContour(contours, current);
                    current = hasCurrentPoint
                        ? new List<OfficePoint> { currentPoint }
                        : null;
                    break;
            }
        }

        AddFlattenedContour(contours, current);
        return contours;
    }

    private static void AddFlattenedContour(List<List<OfficePoint>> contours, List<OfficePoint>? contour) {
        if (contour == null || contour.Count < 3) {
            return;
        }

        if (contour[0].X == contour[contour.Count - 1].X &&
            contour[0].Y == contour[contour.Count - 1].Y) {
            contour.RemoveAt(contour.Count - 1);
        }

        if (contour.Count >= 3) {
            contours.Add(contour);
        }
    }

    private static List<OfficePoint> ClipPolygonToConvexPolygon(IReadOnlyList<OfficePoint> subject, List<OfficePoint> clip) {
        var output = new List<OfficePoint>(subject);
        if (clip.Count < 3) {
            output.Clear();
            return output;
        }

        bool positiveArea = SignedArea(clip) >= 0D;
        for (int i = 0; i < clip.Count && output.Count > 0; i++) {
            OfficePoint edgeStart = clip[i];
            OfficePoint edgeEnd = clip[(i + 1) % clip.Count];
            var input = output;
            output = new List<OfficePoint>();
            OfficePoint previous = input[input.Count - 1];
            bool previousInside = IsInsideClipEdge(previous, edgeStart, edgeEnd, positiveArea);
            for (int j = 0; j < input.Count; j++) {
                OfficePoint current = input[j];
                bool currentInside = IsInsideClipEdge(current, edgeStart, edgeEnd, positiveArea);
                if (currentInside) {
                    if (!previousInside) {
                        output.Add(IntersectLines(previous, current, edgeStart, edgeEnd));
                    }

                    output.Add(current);
                } else if (previousInside) {
                    output.Add(IntersectLines(previous, current, edgeStart, edgeEnd));
                }

                previous = current;
                previousInside = currentInside;
            }
        }

        return output;
    }

    private static bool IsInsideClipEdge(OfficePoint point, OfficePoint edgeStart, OfficePoint edgeEnd, bool positiveArea) {
        double cross = ((edgeEnd.X - edgeStart.X) * (point.Y - edgeStart.Y)) -
            ((edgeEnd.Y - edgeStart.Y) * (point.X - edgeStart.X));
        return positiveArea ? cross >= 0D : cross <= 0D;
    }

    private static OfficePoint IntersectLines(OfficePoint firstStart, OfficePoint firstEnd, OfficePoint secondStart, OfficePoint secondEnd) {
        double x1 = firstStart.X;
        double y1 = firstStart.Y;
        double x2 = firstEnd.X;
        double y2 = firstEnd.Y;
        double x3 = secondStart.X;
        double y3 = secondStart.Y;
        double x4 = secondEnd.X;
        double y4 = secondEnd.Y;
        double denominator = ((x1 - x2) * (y3 - y4)) - ((y1 - y2) * (x3 - x4));
        if (denominator == 0D) {
            return firstEnd;
        }

        double px = ((((x1 * y2) - (y1 * x2)) * (x3 - x4)) - ((x1 - x2) * ((x3 * y4) - (y3 * x4)))) / denominator;
        double py = ((((x1 * y2) - (y1 * x2)) * (y3 - y4)) - ((y1 - y2) * ((x3 * y4) - (y3 * x4)))) / denominator;
        return new OfficePoint(px, py);
    }

    private static double SignedArea(List<OfficePoint> contour) {
        double area = 0D;
        for (int i = 0; i < contour.Count; i++) {
            OfficePoint current = contour[i];
            OfficePoint next = contour[(i + 1) % contour.Count];
            area += (current.X * next.Y) - (next.X * current.Y);
        }

        return area / 2D;
    }

    private static List<OfficePathCommand> BuildClosedContourCommands(List<List<OfficePoint>> contours) {
        var commands = new List<OfficePathCommand>();
        for (int i = 0; i < contours.Count; i++) {
            List<OfficePoint> contour = contours[i];
            if (contour.Count < 3) {
                continue;
            }

            commands.Add(OfficePathCommand.MoveTo(contour[0].X, contour[0].Y));
            for (int j = 1; j < contour.Count; j++) {
                if (!NearlyEqual(contour[j].X, contour[j - 1].X) || !NearlyEqual(contour[j].Y, contour[j - 1].Y)) {
                    commands.Add(OfficePathCommand.LineTo(contour[j].X, contour[j].Y));
                }
            }

            commands.Add(OfficePathCommand.Close());
        }

        return commands;
    }

    private static List<OfficePathCommand> ClipPathCommandsToRectangle(IReadOnlyList<OfficePathCommand> commands, PdfPageClipPath rectangle) {
        var clippedCommands = new List<OfficePathCommand>();
        List<OfficePoint>? current = null;
        OfficePoint currentPoint = default;
        bool hasCurrentPoint = false;
        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    AddClippedContour(clippedCommands, current, rectangle);
                    currentPoint = command.Point;
                    current = new List<OfficePoint> { currentPoint };
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.LineTo:
                    EnsureContour(ref current, currentPoint, hasCurrentPoint);
                    currentPoint = command.Point;
                    current!.Add(currentPoint);
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    EnsureContour(ref current, currentPoint, hasCurrentPoint);
                    current!.AddRange(OfficeGeometry.CreateQuadraticBezierPoints(currentPoint, command.ControlPoint1, command.Point, 24));
                    currentPoint = command.Point;
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    EnsureContour(ref current, currentPoint, hasCurrentPoint);
                    current!.AddRange(OfficeGeometry.CreateCubicBezierPoints(currentPoint, command.ControlPoint1, command.ControlPoint2, command.Point, 24));
                    currentPoint = command.Point;
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.Close:
                    if (current != null && current.Count > 0) {
                        currentPoint = current[0];
                        hasCurrentPoint = true;
                    }
                    AddClippedContour(clippedCommands, current, rectangle);
                    current = hasCurrentPoint
                        ? new List<OfficePoint> { currentPoint }
                        : null;
                    break;
            }
        }

        AddClippedContour(clippedCommands, current, rectangle);
        return clippedCommands;
    }

    private static void EnsureContour(ref List<OfficePoint>? current, OfficePoint currentPoint, bool hasCurrentPoint) {
        if (current == null) {
            current = hasCurrentPoint ? new List<OfficePoint> { currentPoint } : new List<OfficePoint>();
        }
    }

    private static void AddClippedContour(List<OfficePathCommand> commands, List<OfficePoint>? contour, PdfPageClipPath rectangle) {
        if (contour == null || contour.Count < 3) {
            return;
        }

        List<OfficePoint> clipped = ClipPolygonToRectangle(contour, rectangle);
        if (clipped.Count < 3) {
            return;
        }

        commands.Add(OfficePathCommand.MoveTo(clipped[0].X, clipped[0].Y));
        for (int i = 1; i < clipped.Count; i++) {
            if (!NearlyEqual(clipped[i].X, clipped[i - 1].X) || !NearlyEqual(clipped[i].Y, clipped[i - 1].Y)) {
                commands.Add(OfficePathCommand.LineTo(clipped[i].X, clipped[i].Y));
            }
        }

        commands.Add(OfficePathCommand.Close());
    }

    private static List<OfficePoint> ClipPolygonToRectangle(IReadOnlyList<OfficePoint> polygon, PdfPageClipPath rectangle) {
        List<OfficePoint> points = new(polygon);
        points = ClipPolygon(points, point => point.X >= rectangle.X, (from, to) => IntersectVertical(from, to, rectangle.X));
        points = ClipPolygon(points, point => point.X <= rectangle.X + rectangle.Width, (from, to) => IntersectVertical(from, to, rectangle.X + rectangle.Width));
        points = ClipPolygon(points, point => point.Y >= rectangle.Y, (from, to) => IntersectHorizontal(from, to, rectangle.Y));
        points = ClipPolygon(points, point => point.Y <= rectangle.Y + rectangle.Height, (from, to) => IntersectHorizontal(from, to, rectangle.Y + rectangle.Height));
        return points;
    }

    private static List<OfficePoint> ClipPolygon(List<OfficePoint> input, Func<OfficePoint, bool> inside, Func<OfficePoint, OfficePoint, OfficePoint> intersect) {
        var output = new List<OfficePoint>();
        if (input.Count == 0) {
            return output;
        }

        OfficePoint previous = input[input.Count - 1];
        bool previousInside = inside(previous);
        for (int i = 0; i < input.Count; i++) {
            OfficePoint current = input[i];
            bool currentInside = inside(current);
            if (currentInside) {
                if (!previousInside) {
                    output.Add(intersect(previous, current));
                }

                output.Add(current);
            } else if (previousInside) {
                output.Add(intersect(previous, current));
            }

            previous = current;
            previousInside = currentInside;
        }

        return output;
    }

    private static OfficePoint IntersectVertical(OfficePoint from, OfficePoint to, double x) {
        double denominator = to.X - from.X;
        double t = Math.Abs(denominator) <= 0.000001D ? 0D : (x - from.X) / denominator;
        return new OfficePoint(x, from.Y + ((to.Y - from.Y) * t));
    }

    private static OfficePoint IntersectHorizontal(OfficePoint from, OfficePoint to, double y) {
        double denominator = to.Y - from.Y;
        double t = Math.Abs(denominator) <= 0.000001D ? 0D : (y - from.Y) / denominator;
        return new OfficePoint(from.X + ((to.X - from.X) * t), y);
    }

    public static bool TryCreatePath(IReadOnlyList<OfficePathCommand> commands, OfficeFillRule fillRule, out PdfPageClipPath clipPath) {
        clipPath = default;
        if (commands.Count == 0 || commands[0].Kind != OfficePathCommandKind.MoveTo) {
            return false;
        }

        bool hasPoint = false;
        bool hasDraw = false;
        double left = 0D;
        double top = 0D;
        double right = 0D;
        double bottom = 0D;
        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    break;
                case OfficePathCommandKind.LineTo:
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    Include(command.ControlPoint1, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    Include(command.ControlPoint1, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.ControlPoint2, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.Close:
                    break;
            }
        }

        double width = right - left;
        double height = bottom - top;
        if (!hasDraw || width <= 0D || height <= 0D) {
            return false;
        }

        clipPath = new PdfPageClipPath(left, top, width, height, false, fillRule, CloseFilledSubpaths(commands));
        return true;
    }

    private static List<OfficePathCommand> CloseFilledSubpaths(IReadOnlyList<OfficePathCommand> commands) {
        var closed = new List<OfficePathCommand>(commands.Count + 4);
        bool hasOpenSubpath = false;
        bool subpathHasDraw = false;
        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            if (command.Kind == OfficePathCommandKind.MoveTo) {
                if (hasOpenSubpath && subpathHasDraw) {
                    closed.Add(OfficePathCommand.Close());
                }

                hasOpenSubpath = true;
                subpathHasDraw = false;
                closed.Add(command);
                continue;
            }

            closed.Add(command);
            if (command.Kind == OfficePathCommandKind.Close) {
                hasOpenSubpath = false;
                subpathHasDraw = false;
            } else if (command.Kind == OfficePathCommandKind.LineTo ||
                command.Kind == OfficePathCommandKind.QuadraticBezierTo ||
                command.Kind == OfficePathCommandKind.CubicBezierTo) {
                subpathHasDraw = true;
            }
        }

        if (hasOpenSubpath && subpathHasDraw) {
            closed.Add(OfficePathCommand.Close());
        }

        return closed;
    }

    public double X { get; }

    public double Y { get; }

    public double Width { get; }

    public double Height { get; }

    public bool IsRectangle { get; }

    public OfficeFillRule FillRule { get; }

    public IReadOnlyList<OfficePathCommand> Commands { get; }

    internal bool IsExact { get; }

    internal bool CanProveExactIntersection => CanServeAsExactPathClip(this);

    internal bool CanProveNoPositiveAreaIntersection(PdfPageClipPath other) {
        if (!IsExact || !other.IsExact || ContainsCurve(Commands) || ContainsCurve(other.Commands)) return false;
        List<List<OfficePoint>> first = GetContours(this);
        List<List<OfficePoint>> second = GetContours(other);
        if (first.Count == 0 || second.Count == 0) return false;
        for (int firstIndex = 0; firstIndex < first.Count; firstIndex++) {
            for (int secondIndex = 0; secondIndex < second.Count; secondIndex++) {
                if (ContoursIntersect(first[firstIndex], second[secondIndex])) return false;
            }
        }
        for (int contourIndex = 0; contourIndex < first.Count; contourIndex++) {
            for (int pointIndex = 0; pointIndex < first[contourIndex].Count; pointIndex++) {
                if (ContainsFilledPoint(second, other.FillRule, first[contourIndex][pointIndex])) return false;
            }
        }
        for (int contourIndex = 0; contourIndex < second.Count; contourIndex++) {
            for (int pointIndex = 0; pointIndex < second[contourIndex].Count; pointIndex++) {
                if (ContainsFilledPoint(first, FillRule, second[contourIndex][pointIndex])) return false;
            }
        }
        return true;
    }

    private static List<List<OfficePoint>> GetContours(PdfPageClipPath path) {
        if (!path.IsRectangle) return FlattenPathContours(path.Commands);
        return new List<List<OfficePoint>> {
            new List<OfficePoint> {
                new OfficePoint(path.X, path.Y),
                new OfficePoint(path.X + path.Width, path.Y),
                new OfficePoint(path.X + path.Width, path.Y + path.Height),
                new OfficePoint(path.X, path.Y + path.Height)
            }
        };
    }

    private static bool ContoursIntersect(List<OfficePoint> first, List<OfficePoint> second) {
        for (int firstIndex = 0; firstIndex < first.Count; firstIndex++) {
            OfficePoint firstStart = first[firstIndex];
            OfficePoint firstEnd = first[(firstIndex + 1) % first.Count];
            for (int secondIndex = 0; secondIndex < second.Count; secondIndex++) {
                OfficePoint secondStart = second[secondIndex];
                OfficePoint secondEnd = second[(secondIndex + 1) % second.Count];
                if (SegmentsIntersect(firstStart, firstEnd, secondStart, secondEnd)) return true;
            }
        }
        return false;
    }

    private static bool SegmentsIntersect(OfficePoint firstStart, OfficePoint firstEnd, OfficePoint secondStart, OfficePoint secondEnd) {
        double firstA = Cross(firstStart, firstEnd, secondStart);
        double firstB = Cross(firstStart, firstEnd, secondEnd);
        double secondA = Cross(secondStart, secondEnd, firstStart);
        double secondB = Cross(secondStart, secondEnd, firstEnd);
        if (Math.Abs(firstA) <= 0.001D && IsPointOnSegment(secondStart, firstStart, firstEnd) ||
            Math.Abs(firstB) <= 0.001D && IsPointOnSegment(secondEnd, firstStart, firstEnd) ||
            Math.Abs(secondA) <= 0.001D && IsPointOnSegment(firstStart, secondStart, secondEnd) ||
            Math.Abs(secondB) <= 0.001D && IsPointOnSegment(firstEnd, secondStart, secondEnd)) return true;
        return Math.Sign(firstA) != Math.Sign(firstB) && Math.Sign(secondA) != Math.Sign(secondB);
    }

    private static double Cross(OfficePoint start, OfficePoint end, OfficePoint point) =>
        ((end.X - start.X) * (point.Y - start.Y)) - ((end.Y - start.Y) * (point.X - start.X));

    private PdfPageClipPath WithExactness(bool isExact) =>
        IsExact == isExact
            ? this
            : new PdfPageClipPath(X, Y, Width, Height, IsRectangle, FillRule, Commands, isExact);

    internal PdfPageClipPath WithBounds(PdfPageClipPath bounds) {
        if (IsRectangle) {
            return new PdfPageClipPath(bounds.X, bounds.Y, bounds.Width, bounds.Height, true, FillRule, Commands, IsExact);
        }

        List<OfficePathCommand> clippedCommands = ClipPathCommandsToRectangle(Commands, bounds);
        PdfPageClipPath result = clippedCommands.Count > 0 && TryCreatePath(clippedCommands, FillRule, out PdfPageClipPath clippedPath)
            ? clippedPath
            : Rectangle(bounds.X, bounds.Y, 0D, 0D);
        return result.WithExactness(IsExact &&
            !ContainsCurve(Commands) &&
            HasRepresentableClippedContours(this, result));
    }

    internal PdfPageClipPath Translate(double offsetX, double offsetY) {
        if (IsRectangle) {
            return Rectangle(X - offsetX, Y - offsetY, Width, Height).WithExactness(IsExact);
        }

        var translated = new List<OfficePathCommand>(Commands.Count);
        for (int i = 0; i < Commands.Count; i++) {
            OfficePathCommand command = Commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    translated.Add(OfficePathCommand.MoveTo(command.Point.X - offsetX, command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.LineTo:
                    translated.Add(OfficePathCommand.LineTo(command.Point.X - offsetX, command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    translated.Add(OfficePathCommand.QuadraticBezierTo(
                        command.ControlPoint1.X - offsetX,
                        command.ControlPoint1.Y - offsetY,
                        command.Point.X - offsetX,
                        command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    translated.Add(OfficePathCommand.CubicBezierTo(
                        command.ControlPoint1.X - offsetX,
                        command.ControlPoint1.Y - offsetY,
                        command.ControlPoint2.X - offsetX,
                        command.ControlPoint2.Y - offsetY,
                        command.Point.X - offsetX,
                        command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.Close:
                    translated.Add(OfficePathCommand.Close());
                    break;
            }
        }

        return new PdfPageClipPath(X - offsetX, Y - offsetY, Width, Height, false, FillRule, translated, IsExact);
    }

    public OfficeClipPath? ToOfficeClipPath(double primitiveX, double primitiveY) {
        if (!NearlyEqual(X, primitiveX) || !NearlyEqual(Y, primitiveY) || Width <= 0D || Height <= 0D) {
            return null;
        }

        if (IsRectangle) {
            return OfficeClipPath.Rectangle(Width, Height);
        }

        try {
            return OfficeClipPath.Path(TranslateCommands(primitiveX, primitiveY), FillRule);
        } catch (ArgumentException) {
            return null;
        }
    }

    private List<OfficePathCommand> TranslateCommands(double offsetX, double offsetY) {
        var result = new List<OfficePathCommand>(Commands.Count);
        for (int i = 0; i < Commands.Count; i++) {
            OfficePathCommand command = Commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    result.Add(OfficePathCommand.MoveTo(command.Point.X - offsetX, command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.LineTo:
                    result.Add(OfficePathCommand.LineTo(command.Point.X - offsetX, command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    result.Add(OfficePathCommand.QuadraticBezierTo(
                        command.ControlPoint1.X - offsetX,
                        command.ControlPoint1.Y - offsetY,
                        command.Point.X - offsetX,
                        command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    result.Add(OfficePathCommand.CubicBezierTo(
                        command.ControlPoint1.X - offsetX,
                        command.ControlPoint1.Y - offsetY,
                        command.ControlPoint2.X - offsetX,
                        command.ControlPoint2.Y - offsetY,
                        command.Point.X - offsetX,
                        command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.Close:
                    result.Add(OfficePathCommand.Close());
                    break;
            }
        }

        return result;
    }

    private static void Include(OfficePoint point, ref bool hasPoint, ref double left, ref double top, ref double right, ref double bottom) {
        if (!hasPoint) {
            left = right = point.X;
            top = bottom = point.Y;
            hasPoint = true;
            return;
        }

        if (point.X < left) {
            left = point.X;
        }

        if (point.Y < top) {
            top = point.Y;
        }

        if (point.X > right) {
            right = point.X;
        }

        if (point.Y > bottom) {
            bottom = point.Y;
        }
    }

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
}
