using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using UglyToad.PdfPig.Content;

namespace OfficeIMO.ConversionConsistency;

internal static class SvgTextPositions {
    // The owned exporters retain logical labels on glyph-path groups. Ink bounds
    // are not baselines: use the independent saved-PDF glyph box's offset from
    // its baseline, then check the same authored origin and tolerance as text.
    internal static List<TextPositionExpectation> CheckOutlines(XDocument svg, Page pdf,
        PageExpectation expected, int dpi, List<string> errors) {
        XElement[] carriers = svg.Descendants().Where(element => element.Name.LocalName == "text" ||
            element.Attribute("aria-label") != null && !element.Descendants().Any(child =>
                child.Name.LocalName == "text" || child.Attribute("aria-label") != null)).ToArray();
        string[] labels = carriers.Select(Label).ToArray();
        string logical = string.Concat(labels);
        string pdfText = string.Concat(pdf.Letters.Select(letter => letter.Value));
        var selectable = new List<TextPositionExpectation>();
        foreach (TextPositionExpectation position in expected.TextPositions.Concat(expected.SvgTextPositions)) {
            int offset = logical.IndexOf(position.Text, StringComparison.Ordinal);
            int carrierIndex = 0, carrierStart = 0;
            while (carrierIndex < carriers.Length && carrierStart + labels[carrierIndex].Length <= offset)
                carrierStart += labels[carrierIndex++].Length;
            if (offset < 0 || carrierIndex == carriers.Length) {
                errors.Add("SVG is missing positioned text: " + position.Text);
                continue;
            }
            XElement carrier = carriers[carrierIndex];
            if (carrier.Name.LocalName == "text") {
                selectable.Add(position);
                continue;
            }
            string label = labels[carrierIndex];
            int pdfPositionOffset = pdfText.IndexOf(position.Text, StringComparison.Ordinal);
            int withinCarrier = offset - carrierStart;
            int pdfOffset = pdfPositionOffset - withinCarrier;
            if (pdfPositionOffset < 0 || pdfOffset < 0 || pdfOffset + label.Length > pdfText.Length ||
                !pdfText.AsSpan(pdfOffset, label.Length).SequenceEqual(label.AsSpan()) ||
                pdfText.IndexOf(position.Text, pdfPositionOffset + position.Text.Length, StringComparison.Ordinal) >= 0 ||
                logical.IndexOf(position.Text, offset + position.Text.Length, StringComparison.Ordinal) >= 0) {
                errors.Add("SVG outlined text " + position.Text + ": no unambiguous matching saved-PDF glyph oracle for " + label + ".");
                continue;
            }
            Letter[] letters = LettersAt(pdf, pdfOffset, label.Length).ToArray();
            Letter? origin = LettersAt(pdf, pdfPositionOffset, 1).FirstOrDefault();
            Letter[] ink = letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
            if (origin == null || ink.Length == 0 || !TryReadInk(svg, carrier, expected, out InkBounds actual)) {
                errors.Add("SVG outlined text " + position.Text + ": missing finite painted glyph geometry.");
                continue;
            }
            double scale = dpi / 72D;
            var reference = new InkBounds(ink.Min(letter => letter.BoundingBox.Left) * scale,
                ((double)pdf.Height - ink.Max(letter => letter.BoundingBox.Top)) * scale,
                ink.Max(letter => letter.BoundingBox.Right) * scale,
                ((double)pdf.Height - ink.Min(letter => letter.BoundingBox.Bottom)) * scale);
            double x = actual.Left - (reference.Left - origin.StartBaseLine.X * scale);
            double baseline = actual.Top - (reference.Top - ((double)pdf.Height - origin.StartBaseLine.Y) * scale);
            if (!reference.IsFinite || reference.Width <= 0 || reference.Height <= 0 || !double.IsFinite(x) || !double.IsFinite(baseline) ||
                Math.Abs(x - position.X) > position.Tolerance || Math.Abs(baseline - position.Baseline) > position.Tolerance ||
                Math.Abs(actual.Width - reference.Width) > position.Tolerance || Math.Abs(actual.Height - reference.Height) > position.Tolerance)
                errors.Add($"SVG outlined text {position.Text}: baseline ({x:F3},{baseline:F3}), ink {actual.Width:F3}x{actual.Height:F3}; " +
                    $"expected ({position.X},{position.Baseline}), saved-PDF ink {reference.Width:F3}x{reference.Height:F3} within {position.Tolerance}px.");
        }
        return selectable;
    }

    private static string Label(XElement element) => element.Attribute("aria-label")?.Value ?? element.Value;

    private static IEnumerable<Letter> LettersAt(Page page, int offset, int length) {
        int cursor = 0;
        foreach (Letter letter in page.Letters) {
            int end = cursor + letter.Value.Length;
            if (cursor < offset + length && end > offset) yield return letter;
            cursor = end;
        }
    }

    private static bool TryReadInk(XDocument source, XElement carrier, PageExpectation expected, out InkBounds bounds) {
        bounds = default;
        if (source.Root == null) return false;
        // Preserve inherited styles, transforms, definitions and viewport mapping;
        // the canonical reader owns SVG/path parsing. This projection never alters
        // the saved artifact or supplies product-only identifiers/metadata.
        XElement branch = new(carrier);
        foreach (XElement ancestor in carrier.Ancestors().TakeWhile(element => element != source.Root))
            branch = new XElement(ancestor.Name, ancestor.Attributes(), Definitions(ancestor), branch);
        var isolated = new XElement(source.Root.Name, source.Root.Attributes(), Definitions(source.Root), branch);
        var options = new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 100000 };
        if (!OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(isolated.ToString(SaveOptions.DisableFormatting)),
                options, out OfficeDrawing? drawing, out int unsupported) || drawing == null || unsupported != 0) return false;
        var points = new List<OfficePoint>();
        if (!ReadPoints(drawing, OfficeTransform.Identity, points) || points.Count < 3) return false;
        double sx = expected.Width / drawing.Width, sy = expected.Height / drawing.Height;
        bounds = new InkBounds(points.Min(point => point.X) * sx, points.Min(point => point.Y) * sy,
            points.Max(point => point.X) * sx, points.Max(point => point.Y) * sy);
        return bounds.IsFinite && bounds.Width > 0 && bounds.Height > 0;
    }

    private static IEnumerable<XElement> Definitions(XElement element) => element.Elements()
        .Where(child => child.Name.LocalName is "defs" or "style").Select(child => new XElement(child));

    private static bool ReadPoints(OfficeDrawing drawing, OfficeTransform parent, List<OfficePoint> points) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            switch (element) {
                case OfficeDrawingShape placed:
                    OfficeShape shape = placed.Shape;
                    bool painted = shape.FillColor is OfficeColor fill && fill.A > 0 && shape.FillOpacity != 0 ||
                        shape.StrokeColor is OfficeColor stroke && stroke.A > 0 && shape.StrokeOpacity != 0;
                    if (!painted) continue;
                    OfficeTransform transform = (shape.Transform ?? OfficeTransform.Identity)
                        .Then(OfficeTransform.Translate(placed.X, placed.Y)).Then(parent);
                    foreach (OfficePathCommand command in shape.PathCommands) {
                        if (command.Kind == OfficePathCommandKind.Close) continue;
                        // Owned glyph exports are flattened contours. Other shapes
                        // need their own explicit geometry contract rather than a
                        // control-point approximation of painted curve bounds.
                        if (command.Kind is not (OfficePathCommandKind.MoveTo or OfficePathCommandKind.LineTo)) return false;
                        OfficePoint point = transform.TransformPoint(command.Point);
                        if (!double.IsFinite(point.X) || !double.IsFinite(point.Y)) return false;
                        points.Add(point);
                    }
                    break;
                case OfficeDrawingGroup group:
                    OfficeTransform placement = OfficeTransform.Translate(group.X + group.ContentOffsetX, group.Y + group.ContentOffsetY);
                    if (group.FrameTransform.HasValue) placement = placement.Then(group.FrameTransform.Value.CreateDestinationTransform());
                    if (!ReadPoints(group.Drawing, placement.Then(parent), points)) return false;
                    break;
                case OfficeDrawingEffectGroup effect:
                    if (effect.Opacity > 0 && !ReadPoints(effect.Drawing, effect.Transform.Then(parent), points)) return false;
                    break;
                default:
                    return false;
            }
        }
        return true;
    }

    private readonly record struct InkBounds(double Left, double Top, double Right, double Bottom) {
        internal double Width => Right - Left;
        internal double Height => Bottom - Top;
        internal bool IsFinite => double.IsFinite(Left) && double.IsFinite(Top) && double.IsFinite(Right) && double.IsFinite(Bottom);
    }
}
