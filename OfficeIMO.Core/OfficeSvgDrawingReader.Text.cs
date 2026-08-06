using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumTextRuns = 4096;

    private static void AddText(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext style,
        SvgPaintServerRegistry paintServers,
        OfficeTransform transform,
        double viewX,
        double viewY,
        ref int unsupported) {
        var runs = new List<SvgTextRun>();
        var cursor = new SvgTextCursor { Chunk = -1 };
        bool preserve = string.Equals(element.Attribute(XNamespace.Xml + "space")?.Value, "preserve", StringComparison.OrdinalIgnoreCase);
        AddTextElementRuns(element, style, paintServers, transform, preserve, false, viewX, viewY,
            drawing.Width, drawing.Height, runs, 0, ref cursor, ref unsupported);
        if (runs.Count == 0) return;
        ApplyTextAnchors(runs);
        foreach (SvgTextRun run in runs) AddTextRun(drawing, run, ref unsupported);
    }

    private static void AddTextElementRuns(
        XElement element,
        SvgPaintContext inheritedStyle,
        SvgPaintServerRegistry paintServers,
        OfficeTransform inheritedTransform,
        bool inheritedPreserve,
        bool resolveElement,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight,
        ICollection<SvgTextRun> runs,
        int depth,
        ref SvgTextCursor cursor,
        ref int unsupported) {
        if (depth > MaximumSvgNestingDepth) {
            unsupported++;
            return;
        }
        if (runs.Count >= MaximumTextRuns) {
            ReportTextRunLimit(ref cursor, ref unsupported);
            return;
        }

        SvgPaintContext style = resolveElement
            ? ResolvePaintContext(element, inheritedStyle, paintServers, ref unsupported)
            : inheritedStyle;
        if (!style.Visible) return;
        OfficeTransform transform = resolveElement
            ? ResolveTransform(element, inheritedTransform, viewX, viewY, ref unsupported)
            : inheritedTransform;
        bool preserve = inheritedPreserve;
        string? space = element.Attribute(XNamespace.Xml + "space")?.Value;
        if (!string.IsNullOrWhiteSpace(space)) {
            if (space!.Equals("preserve", StringComparison.OrdinalIgnoreCase)) preserve = true;
            else if (space.Equals("default", StringComparison.OrdinalIgnoreCase)) preserve = false;
            else unsupported++;
        }
        ApplyTextPosition(element, viewX, viewY, viewportWidth, viewportHeight, ref cursor, ref unsupported);
        int firstRun = runs.Count;
        double lengthOrigin = cursor.X;
        bool adjustGlyphs = TryReadTextLengthAdjustment(element, out double authoredLength, ref unsupported);

        foreach (XNode node in element.Nodes()) {
            if (runs.Count >= MaximumTextRuns) {
                ReportTextRunLimit(ref cursor, ref unsupported);
                return;
            }
            if (node is XText textNode) {
                string text = NormalizeText(textNode.Value, preserve, ref cursor);
                if (text.Length == 0) continue;
                double fontSize = Math.Max(0.1D, style.FontSize);
                double width = Math.Max(0.1D, text.Length * fontSize * 0.62D);
                double baseline = ResolveTextBaseline(cursor.Baseline, fontSize, style.DominantBaseline);
                runs.Add(new SvgTextRun(text, cursor.X, baseline, width, fontSize, cursor.Chunk, style.TextAnchor, style, transform));
                cursor.X += width;
                cursor.HasText = true;
                continue;
            }
            if (node is XElement child && child.Name.LocalName.Equals("tspan", StringComparison.OrdinalIgnoreCase)) {
                AddTextElementRuns(child, style, paintServers, transform, preserve, true, viewX, viewY,
                    viewportWidth, viewportHeight, runs, depth + 1, ref cursor, ref unsupported);
            } else if (node is XElement) {
                unsupported++;
            }
        }
        if (adjustGlyphs) ApplyTextLengthAdjustment(runs, firstRun, lengthOrigin, authoredLength, ref cursor, ref unsupported);
    }

    private static bool TryReadTextLengthAdjustment(XElement element, out double textLength, ref int unsupported) {
        textLength = 0D;
        string? value = element.Attribute("textLength")?.Value;
        string? mode = element.Attribute("lengthAdjust")?.Value;
        if (string.IsNullOrWhiteSpace(value)) {
            if (!string.IsNullOrWhiteSpace(mode)) unsupported++;
            return false;
        }
        if (!TrySvgLength(value, out textLength) || textLength <= 0D) {
            unsupported++;
            return false;
        }
        if (string.IsNullOrWhiteSpace(mode) || !mode!.Trim().Equals("spacingAndGlyphs", StringComparison.OrdinalIgnoreCase)) {
            unsupported++;
            return false;
        }
        return true;
    }

    private static void ApplyTextLengthAdjustment(
        ICollection<SvgTextRun> runs,
        int firstRun,
        double origin,
        double authoredLength,
        ref SvgTextCursor cursor,
        ref int unsupported) {
        if (firstRun >= runs.Count) return;
        SvgTextRun[] adjusted = runs.Skip(firstRun).ToArray();
        double right = adjusted.Max(run => run.X + run.Width);
        double naturalLength = right - origin;
        if (naturalLength <= 0.0000001D) {
            unsupported++;
            return;
        }
        double scale = authoredLength / naturalLength;
        if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0.01D || scale > 100D) {
            unsupported++;
            return;
        }
        foreach (SvgTextRun run in adjusted) {
            run.X = origin + ((run.X - origin) * scale);
            run.Width *= scale;
            run.GlyphScale *= scale;
        }
        cursor.X = origin + ((cursor.X - origin) * scale);
    }

    private static void ApplyTextPosition(
        XElement element,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight,
        ref SvgTextCursor cursor,
        ref int unsupported) {
        bool startsChunk = false;
        if (TryTextLength(element, "x", viewportWidth, out double x, out _, ref unsupported)) {
            cursor.X = x - viewX;
            startsChunk = true;
        }
        if (TryTextLength(element, "y", viewportHeight, out double y, out _, ref unsupported)) {
            cursor.Baseline = y - viewY;
            startsChunk = true;
        }
        if (cursor.Chunk < 0 || startsChunk) {
            cursor.Chunk++;
            cursor.PendingSpace = false;
        }
        if (TryTextLength(element, "dx", viewportWidth, out double dx, out _, ref unsupported)) cursor.X += dx;
        if (TryTextLength(element, "dy", viewportHeight, out double dy, out _, ref unsupported)) cursor.Baseline += dy;
    }

    private static bool TryTextLength(
        XElement element,
        string name,
        double percentageReference,
        out double value,
        out bool percentage,
        ref int unsupported) {
        value = 0D;
        percentage = false;
        string? text = element.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(text)) return false;
        string[] values = text!.Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (values.Length != 1) unsupported++;
        if (values.Length > 0 && TryViewportLength(values[0], percentageReference, out value, out percentage)) return true;
        unsupported++;
        return false;
    }

    private static double ResolveTextBaseline(double authoredY, double fontSize, SvgDominantBaseline baseline) {
        switch (baseline) {
            case SvgDominantBaseline.Hanging:
                return authoredY + fontSize;
            case SvgDominantBaseline.Middle:
                return authoredY + (fontSize / 2D);
            case SvgDominantBaseline.TextAfterEdge:
            case SvgDominantBaseline.Alphabetic:
            default:
                return authoredY;
        }
    }

    private static string NormalizeText(string raw, bool preserve, ref SvgTextCursor cursor) {
        if (preserve) {
            cursor.PendingSpace = false;
            return raw;
        }

        var builder = new StringBuilder(raw.Length);
        foreach (char character in raw) {
            if (char.IsWhiteSpace(character)) {
                if (cursor.HasText || builder.Length > 0) cursor.PendingSpace = true;
                continue;
            }
            if (cursor.PendingSpace && (cursor.HasText || builder.Length > 0)) builder.Append(' ');
            builder.Append(character);
            cursor.PendingSpace = false;
        }
        return builder.ToString();
    }

    private static void ReportTextRunLimit(ref SvgTextCursor cursor, ref int unsupported) {
        if (cursor.LimitReported) return;
        cursor.LimitReported = true;
        unsupported++;
    }

    private static void ApplyTextAnchors(IList<SvgTextRun> runs) {
        foreach (IGrouping<int, SvgTextRun> chunk in runs.GroupBy(run => run.Chunk)) {
            SvgTextRun first = chunk.First();
            if (first.Anchor == "start") continue;
            double left = chunk.Min(run => run.X);
            double right = chunk.Max(run => run.X + run.Width);
            double shift = first.Anchor == "middle" ? -(right - left) / 2D : -(right - left);
            foreach (SvgTextRun run in chunk) run.X += shift;
        }
    }

    private static void AddTextRun(OfficeDrawing drawing, SvgTextRun run, ref int unsupported) {
        if (run.Style.FillGradient != null || run.Style.FillRadialGradient != null || run.Style.FillDeferredGradient != null) {
            unsupported++;
            return;
        }
        if (!run.Style.Fill.HasValue) return;
        double x = run.X;
        double y = run.Baseline - run.FontSize;
        double width = run.Width;
        double height = run.FontSize * 1.25D;
        if (x < 0D || y < 0D || x >= drawing.Width || y >= drawing.Height) {
            unsupported++;
            return;
        }
        width = Math.Min(width, drawing.Width - x);
        height = Math.Min(height, drawing.Height - y);
        if (width <= 0D || height <= 0D) return;

        OfficeColor baseColor = run.Style.Fill.Value;
        double opacity = Math.Max(0D, Math.Min(1D, run.Style.FillOpacity * run.Style.Opacity));
        OfficeColor color = OfficeColor.FromRgba(baseColor.R, baseColor.G, baseColor.B, (byte)Math.Round(baseColor.A * opacity));
        var font = new OfficeFontInfo(run.Style.FontFamily, run.FontSize, run.Style.FontStyle);
        bool usesEffect = run.Transform != OfficeTransform.Identity || Math.Abs(run.GlyphScale - 1D) > 0.0000001D;
        OfficeDrawing target = usesEffect ? new OfficeDrawing(drawing.Width, drawing.Height) : drawing;
        try {
            double naturalWidth = width / run.GlyphScale;
            target.AddText(run.Text, x, y, naturalWidth, height, font, color, OfficeTextAlignment.Left, height);
            if (!ReferenceEquals(target, drawing)) {
                OfficeTransform effect = run.GlyphScale.Equals(1D)
                    ? run.Transform
                    : OfficeTransform.Translate(-x, 0D)
                        .Then(OfficeTransform.Scale(run.GlyphScale, 1D))
                        .Then(OfficeTransform.Translate(x, 0D))
                        .Then(run.Transform);
                drawing.AddEffectDrawing(target, effect);
            }
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
        }
    }

    private sealed class SvgTextRun {
        internal string Text { get; }
        internal double X { get; set; }
        internal double Baseline { get; }
        internal double Width { get; set; }
        internal double GlyphScale { get; set; } = 1D;
        internal double FontSize { get; }
        internal int Chunk { get; }
        internal string Anchor { get; }
        internal SvgPaintContext Style { get; }
        internal OfficeTransform Transform { get; }

        internal SvgTextRun(string text, double x, double baseline, double width, double fontSize, int chunk, string anchor, SvgPaintContext style, OfficeTransform transform) {
            Text = text;
            X = x;
            Baseline = baseline;
            Width = width;
            FontSize = fontSize;
            Chunk = chunk;
            Anchor = anchor;
            Style = style;
            Transform = transform;
        }
    }

    private struct SvgTextCursor {
        internal double X;
        internal double Baseline;
        internal int Chunk;
        internal bool HasText;
        internal bool PendingSpace;
        internal bool LimitReported;
    }
}
