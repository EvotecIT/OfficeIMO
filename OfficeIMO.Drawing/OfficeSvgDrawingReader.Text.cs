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
        AddTextElementRuns(element, style, paintServers, transform, preserve, false, viewX, viewY, runs, ref cursor, ref unsupported);
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
        ICollection<SvgTextRun> runs,
        ref SvgTextCursor cursor,
        ref int unsupported) {
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
        if (element.Attribute("textLength") != null || element.Attribute("lengthAdjust") != null) unsupported++;
        ApplyTextPosition(element, viewX, viewY, ref cursor, ref unsupported);

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
                runs.Add(new SvgTextRun(text, cursor.X, cursor.Baseline, width, fontSize, cursor.Chunk, style.TextAnchor, style, transform));
                cursor.X += width;
                cursor.HasText = true;
                continue;
            }
            if (node is XElement child && child.Name.LocalName.Equals("tspan", StringComparison.OrdinalIgnoreCase)) {
                AddTextElementRuns(child, style, paintServers, transform, preserve, true, viewX, viewY, runs, ref cursor, ref unsupported);
            } else if (node is XElement) {
                unsupported++;
            }
        }
    }

    private static void ApplyTextPosition(XElement element, double viewX, double viewY, ref SvgTextCursor cursor, ref int unsupported) {
        bool startsChunk = false;
        if (TryTextLength(element, "x", out double x, ref unsupported)) {
            cursor.X = x - viewX;
            startsChunk = true;
        }
        if (TryTextLength(element, "y", out double y, ref unsupported)) {
            cursor.Baseline = y - viewY;
            startsChunk = true;
        }
        if (cursor.Chunk < 0 || startsChunk) {
            cursor.Chunk++;
            cursor.PendingSpace = false;
        }
        if (TryTextLength(element, "dx", out double dx, ref unsupported)) cursor.X += dx;
        if (TryTextLength(element, "dy", out double dy, ref unsupported)) cursor.Baseline += dy;
    }

    private static bool TryTextLength(XElement element, string name, out double value, ref int unsupported) {
        value = 0D;
        string? text = element.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(text)) return false;
        string[] values = text!.Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (values.Length != 1) unsupported++;
        if (values.Length > 0 && TrySvgLength(values[0], out value)) return true;
        unsupported++;
        return false;
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
        OfficeDrawing target = run.Transform == OfficeTransform.Identity ? drawing : new OfficeDrawing(drawing.Width, drawing.Height);
        try {
            target.AddText(run.Text, x, y, width, height, font, color, OfficeTextAlignment.Left, height);
            if (!ReferenceEquals(target, drawing)) drawing.AddEffectDrawing(target, run.Transform);
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
        }
    }

    private sealed class SvgTextRun {
        internal string Text { get; }
        internal double X { get; set; }
        internal double Baseline { get; }
        internal double Width { get; }
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
