using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static void ApplyTextPaths(
        IList<SvgTextRun> runs,
        IEnumerable<SvgTextPathLayout> layouts,
        SvgElementReferenceRegistry references,
        double viewX,
        double viewY,
        ref int unsupported) {
        foreach (SvgTextPathLayout layout in layouts.OrderByDescending(item => item.FirstRun)) {
            int end = Math.Min(layout.EndRun, runs.Count);
            if (layout.FirstRun < 0 || layout.FirstRun >= end) continue;
            if (!TryResolveTextPath(layout.Element, references, viewX, viewY, ref unsupported,
                    out SvgFlattenedTextPath? path, out double startOffset)) {
                RemoveTextPathRuns(runs, layout.FirstRun, end);
                continue;
            }

            var replacements = new List<SvgTextRun>();
            for (int runIndex = layout.FirstRun; runIndex < end; runIndex++) {
                SvgTextRun source = runs[runIndex];
                IReadOnlyList<string> glyphs = OfficeTextElements.Split(source.Text);
                if (glyphs.Count == 0) continue;
                double glyphAdvance = source.Width / glyphs.Count;
                for (int glyphIndex = 0; glyphIndex < glyphs.Count; glyphIndex++) {
                    double distance = startOffset + ((source.X + ((glyphIndex + 0.5D) * glyphAdvance)) * path!.AuthorUnitScale);
                    if (!path!.TryResolve(distance, out OfficePoint point, out double angleDegrees)) continue;
                    double glyphWidth = Math.Max(0.1D, glyphAdvance / source.GlyphScale);
                    replacements.Add(new SvgTextRun(
                        glyphs[glyphIndex],
                        point.X - glyphWidth / 2D,
                        point.Y,
                        glyphAdvance,
                        source.FontSize,
                        source.Chunk,
                        "start",
                        source.Style,
                        source.Transform,
                        source.FontProgram,
                        angleDegrees,
                        point.X,
                        point.Y) {
                        GlyphScale = source.GlyphScale
                    });
                }
            }

            RemoveTextPathRuns(runs, layout.FirstRun, end);
            for (int index = 0; index < replacements.Count; index++) {
                runs.Insert(layout.FirstRun + index, replacements[index]);
            }
        }
    }

    private static void RemoveTextPathRuns(IList<SvgTextRun> runs, int first, int end) {
        for (int index = end - 1; index >= first; index--) runs.RemoveAt(index);
    }

    private static bool TryResolveTextPath(
        XElement textPath,
        SvgElementReferenceRegistry references,
        double viewX,
        double viewY,
        ref int unsupported,
        out SvgFlattenedTextPath? path,
        out double startOffset) {
        path = null;
        startOffset = 0D;
        SvgElementReferenceEntryResult entry = references.TryEnterDetailed(
            textPath,
            "path",
            out string referenceId,
            out XElement? target);
        if (entry != SvgElementReferenceEntryResult.Entered) {
            unsupported++;
            return false;
        }

        try {
            if (!OfficeSvgPathDataParser.TryParse(target!.Attribute("d")?.Value, MaximumSvgPathCommands,
                    out IReadOnlyList<OfficePathCommand> parsed, out bool limitExceeded)
                || limitExceeded
                || parsed.Count == 0) {
                unsupported++;
                return false;
            }

            var commands = new List<OfficePathCommand>(parsed.Count);
            foreach (OfficePathCommand command in parsed) commands.Add(command.Translate(viewX, viewY));
            OfficeTransform pathTransform = ResolveTransform(target, OfficeTransform.Identity, viewX, viewY, ref unsupported);
            IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(commands, 0D, 0D, 1D);
            var segments = new List<SvgTextPathSegment>();
            double totalLength = 0D;
            foreach (OfficeFlattenedPathContour contour in contours) {
                AppendTextPathSegments(contour.Points, pathTransform, ref totalLength, segments);
                if (contour.Closed && contour.Points.Count > 1) {
                    AppendTextPathSegment(contour.Points[contour.Points.Count - 1], contour.Points[0], pathTransform,
                        ref totalLength, segments);
                }
            }
            if (segments.Count == 0 || totalLength <= 0.000001D) {
                unsupported++;
                return false;
            }

            string? method = textPath.Attribute("method")?.Value;
            if (!string.IsNullOrWhiteSpace(method) && !method!.Trim().Equals("align", StringComparison.OrdinalIgnoreCase)) unsupported++;
            string? spacing = textPath.Attribute("spacing")?.Value;
            if (!string.IsNullOrWhiteSpace(spacing)
                && !spacing!.Trim().Equals("auto", StringComparison.OrdinalIgnoreCase)
                && !spacing.Trim().Equals("exact", StringComparison.OrdinalIgnoreCase)) unsupported++;

            if (!TryResolveTextPathOffset(textPath.Attribute("startOffset")?.Value, totalLength, out startOffset, out bool percentageOffset)) {
                unsupported++;
                return false;
            }
            double authorUnitScale = 1D;
            if (TrySvgLength(target.Attribute("pathLength")?.Value, out double authoredPathLength) && authoredPathLength > 0D) {
                authorUnitScale = totalLength / authoredPathLength;
                if (!percentageOffset) startOffset *= authorUnitScale;
            }
            path = new SvgFlattenedTextPath(segments, totalLength, authorUnitScale);
            return true;
        } finally {
            references.Exit(referenceId);
        }
    }

    private static void AppendTextPathSegments(
        IReadOnlyList<OfficePoint> points,
        OfficeTransform transform,
        ref double totalLength,
        ICollection<SvgTextPathSegment> segments) {
        for (int index = 1; index < points.Count; index++) {
            AppendTextPathSegment(points[index - 1], points[index], transform, ref totalLength, segments);
        }
    }

    private static void AppendTextPathSegment(
        OfficePoint sourceStart,
        OfficePoint sourceEnd,
        OfficeTransform transform,
        ref double totalLength,
        ICollection<SvgTextPathSegment> segments) {
        OfficePoint start = transform.TransformPoint(sourceStart);
        OfficePoint end = transform.TransformPoint(sourceEnd);
        double dx = end.X - start.X;
        double dy = end.Y - start.Y;
        double length = Math.Sqrt((dx * dx) + (dy * dy));
        if (length <= 0.000001D) return;
        segments.Add(new SvgTextPathSegment(start, end, totalLength, length));
        totalLength += length;
    }

    private static bool TryResolveTextPathOffset(string? value, double totalLength, out double offset, out bool percentage) {
        offset = 0D;
        percentage = false;
        if (string.IsNullOrWhiteSpace(value)) return true;
        string normalized = value!.Trim();
        if (normalized.EndsWith("%", StringComparison.Ordinal)) {
            percentage = true;
            return double.TryParse(normalized.Substring(0, normalized.Length - 1), NumberStyles.Float,
                       CultureInfo.InvariantCulture, out double percentValue)
                   && !double.IsNaN(percentValue)
                   && !double.IsInfinity(percentValue)
                   && (offset = totalLength * percentValue / 100D) == offset;
        }
        return TrySvgLength(normalized, out offset);
    }

    private sealed class SvgFlattenedTextPath {
        private readonly IReadOnlyList<SvgTextPathSegment> _segments;

        internal SvgFlattenedTextPath(IReadOnlyList<SvgTextPathSegment> segments, double length, double authorUnitScale) {
            _segments = segments;
            Length = length;
            AuthorUnitScale = authorUnitScale;
        }

        internal double Length { get; }
        internal double AuthorUnitScale { get; }

        internal bool TryResolve(double distance, out OfficePoint point, out double angleDegrees) {
            point = default;
            angleDegrees = 0D;
            if (distance < 0D || distance > Length) return false;
            SvgTextPathSegment segment = _segments[_segments.Count - 1];
            for (int index = 0; index < _segments.Count; index++) {
                if (distance <= _segments[index].EndDistance) {
                    segment = _segments[index];
                    break;
                }
            }
            double progress = Math.Max(0D, Math.Min(1D, (distance - segment.StartDistance) / segment.Length));
            point = new OfficePoint(
                segment.Start.X + ((segment.End.X - segment.Start.X) * progress),
                segment.Start.Y + ((segment.End.Y - segment.Start.Y) * progress));
            angleDegrees = Math.Atan2(segment.End.Y - segment.Start.Y, segment.End.X - segment.Start.X) * 180D / Math.PI;
            return true;
        }
    }

    private readonly struct SvgTextPathSegment {
        internal SvgTextPathSegment(OfficePoint start, OfficePoint end, double startDistance, double length) {
            Start = start;
            End = end;
            StartDistance = startDistance;
            Length = length;
        }

        internal OfficePoint Start { get; }
        internal OfficePoint End { get; }
        internal double StartDistance { get; }
        internal double Length { get; }
        internal double EndDistance => StartDistance + Length;
    }
}
