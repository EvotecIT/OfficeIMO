using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Bounded arc-length queries and exact Bezier subdivision for one normalized open contour.</summary>
internal sealed class OfficePathMeasure {
    private const int MaximumSamples = 200_000;
    private readonly List<Segment> _segments = new List<Segment>();
    internal double Length { get; }

    internal OfficePathMeasure(IReadOnlyList<OfficePathCommand> commands) {
        int samples = 0;
        Length = Measure(commands, ref samples);
    }

    /// <summary>Shares the sampling budget across contours in one conversion operation.</summary>
    internal OfficePathMeasure(IReadOnlyList<OfficePathCommand> commands, ref int samples) => Length = Measure(commands, ref samples);

    private double Measure(IReadOnlyList<OfficePathCommand> commands, ref int samples) {
        if (!OfficePathGeometry.TryOpenEndpoints(commands, out _, out _, out _, out _))
            throw new NotSupportedException("Path measurement requires one nondegenerate open contour.");
        OfficePoint current = commands[0].Point;
        double length = 0;
        for (int i = 1; i < commands.Count; i++) {
            var segment = new Segment(current, commands[i], length, ref samples);
            if (segment.Length > 0) { _segments.Add(segment); length += segment.Length; }
            if (double.IsInfinity(length)) throw new NotSupportedException("Path length exceeds the finite geometry range.");
            current = commands[i].Point;
        }
        if (!(length > 0)) throw new NotSupportedException("Path has no measurable stroke.");
        return length;
    }

    internal OfficePoint PointAtLength(double distance) {
        Validate(distance);
        distance = Math.Max(0, Math.Min(Length, distance));
        foreach (Segment segment in _segments)
            if (distance <= segment.Offset + segment.Length) return segment.Point(segment.Parameter(distance - segment.Offset));
        return _segments[_segments.Count - 1].Command.Point;
    }

    /// <summary>Clamps distances to the contour; an exhausted interval has no shaft. Curves remain curves.</summary>
    internal IReadOnlyList<OfficePathCommand> Slice(double from, double to) {
        Validate(from); Validate(to);
        from = Math.Max(0, Math.Min(Length, from)); to = Math.Max(0, Math.Min(Length, to));
        var result = new List<OfficePathCommand>();
        if (from >= to) return result;
        foreach (Segment segment in _segments) {
            double first = Math.Max(from, segment.Offset), last = Math.Min(to, segment.Offset + segment.Length);
            if (first >= last) continue;
            double a = segment.Parameter(first - segment.Offset), b = segment.Parameter(last - segment.Offset);
            if (result.Count == 0) result.Add(OfficePathCommand.MoveTo(segment.Point(a)));
            result.Add(segment.Subdivide(a, b));
        }
        return result;
    }

    private static void Validate(double distance) {
        if (double.IsNaN(distance) || double.IsInfinity(distance)) throw new ArgumentOutOfRangeException(nameof(distance));
    }

    private sealed class Segment {
        internal readonly OfficePoint Start;
        internal readonly OfficePathCommand Command;
        internal readonly double Offset, Length;
        private readonly double[] _lengths;

        internal Segment(OfficePoint start, OfficePathCommand command, double offset, ref int samples) {
            Start = start; Command = command; Offset = offset;
            // A finer version of the renderer's bounded sampling estimates arc length only;
            // subdivision below retains the original analytic geometry in the exported path.
            int count = command.Kind == OfficePathCommandKind.LineTo ? 1 : Math.Max(16,
                command.Kind == OfficePathCommandKind.QuadraticBezierTo ?
                    OfficeCurveFlattening.QuadraticSegments(start, command.ControlPoint1, command.Point, 10) :
                    OfficeCurveFlattening.CubicSegments(start, command.ControlPoint1, command.ControlPoint2, command.Point, 10));
            samples += count + 1;
            if (samples > MaximumSamples) throw new NotSupportedException("Path measurement exceeds the sample limit.");
            _lengths = new double[count + 1];
            OfficePoint previous = start;
            for (int i = 1; i <= count; i++) {
                OfficePoint next = Point((double)i/count);
                double x = Math.Abs(next.X - previous.X), y = Math.Abs(next.Y - previous.Y), scale = Math.Max(x, y);
                double length = scale == 0 ? 0 : scale * Math.Sqrt((x/scale)*(x/scale) + (y/scale)*(y/scale));
                if (double.IsNaN(length) || double.IsInfinity(length)) throw new NotSupportedException("Path segment exceeds the finite geometry range.");
                _lengths[i] = _lengths[i - 1] + length; previous = next;
            }
            Length = _lengths[count];
        }

        internal double Parameter(double distance) {
            if (distance <= 0) return 0;
            if (distance >= Length) return 1;
            int index = Array.BinarySearch(_lengths, distance);
            if (index >= 0) return (double)index/(_lengths.Length - 1);
            index = ~index;
            double span = _lengths[index] - _lengths[index - 1];
            return (index - 1 + (distance - _lengths[index - 1])/span)/(_lengths.Length - 1);
        }

        internal OfficePoint Point(double t) {
            if (Command.Kind == OfficePathCommandKind.LineTo) return Mix(Start, Command.Point, t);
            OfficePoint first = Mix(Start, Command.ControlPoint1, t);
            if (Command.Kind == OfficePathCommandKind.QuadraticBezierTo) return Mix(first, Mix(Command.ControlPoint1, Command.Point, t), t);
            OfficePoint second = Mix(Command.ControlPoint1, Command.ControlPoint2, t), third = Mix(Command.ControlPoint2, Command.Point, t);
            return Mix(Mix(first, second, t), Mix(second, third, t), t);
        }

        internal OfficePathCommand Subdivide(double from, double to) {
            if (from == 0 && to == 1) return Command;
            if (Command.Kind == OfficePathCommandKind.LineTo) return OfficePathCommand.LineTo(Point(to));
            // Split at the upper parameter, then retain the right-hand portion of that left curve.
            OfficePoint a = Mix(Start, Command.ControlPoint1, to);
            double relative = from/to;
            if (Command.Kind == OfficePathCommandKind.QuadraticBezierTo) {
                OfficePoint end = Point(to);
                return OfficePathCommand.QuadraticBezierTo(Mix(a, end, relative), end);
            }
            OfficePoint b = Mix(Command.ControlPoint1, Command.ControlPoint2, to);
            OfficePoint left2 = Mix(a, b, to), leftEnd = Point(to);
            OfficePoint right2 = Mix(left2, leftEnd, relative);
            OfficePoint right1 = Mix(Mix(a, left2, relative), right2, relative);
            return OfficePathCommand.CubicBezierTo(right1, right2, leftEnd);
        }

        private static OfficePoint Mix(OfficePoint a, OfficePoint b, double t) => new OfficePoint((1 - t)*a.X + t*b.X, (1 - t)*a.Y + t*b.Y);
    }
}
