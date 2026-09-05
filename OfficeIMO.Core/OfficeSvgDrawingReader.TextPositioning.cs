using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed class SvgTextPositioning {
        private readonly SvgTextPositioning? _parent;
        private readonly IReadOnlyList<double>? _x;
        private readonly IReadOnlyList<double>? _y;
        private readonly IReadOnlyList<double>? _dx;
        private readonly IReadOnlyList<double>? _dy;
        private readonly IReadOnlyList<double>? _rotate;
        private int _index;

        internal bool RequiresPerCharacterRuns =>
            HasRemainingMultipleValues(_x) || HasRemainingMultipleValues(_y)
            || HasRemainingMultipleValues(_dx) || HasRemainingMultipleValues(_dy)
            || HasRemainingMultipleValues(_rotate)
            || _parent?.RequiresPerCharacterRuns == true;

        private SvgTextPositioning(
            SvgTextPositioning? parent,
            IReadOnlyList<double>? x,
            IReadOnlyList<double>? y,
            IReadOnlyList<double>? dx,
            IReadOnlyList<double>? dy,
            IReadOnlyList<double>? rotate) {
            _parent = parent;
            _x = x;
            _y = y;
            _dx = dx;
            _dy = dy;
            _rotate = rotate;
        }

        internal static SvgTextPositioning? Create(
            XElement element,
            SvgTextPositioning? parent,
            double viewportWidth,
            double viewportHeight,
            ref int unsupported) {
            IReadOnlyList<double>? x = ParseLengthList(element, "x", viewportWidth, ref unsupported);
            IReadOnlyList<double>? y = ParseLengthList(element, "y", viewportHeight, ref unsupported);
            IReadOnlyList<double>? dx = ParseLengthList(element, "dx", viewportWidth, ref unsupported);
            IReadOnlyList<double>? dy = ParseLengthList(element, "dy", viewportHeight, ref unsupported);
            IReadOnlyList<double>? rotate = ParseRotationList(element, ref unsupported);
            return x == null && y == null && dx == null && dy == null && rotate == null
                ? parent
                : new SvgTextPositioning(parent, x, y, dx, dy, rotate);
        }

        internal double Apply(ref SvgTextCursor cursor, double viewX, double viewY) {
            bool hasX = TryResolve(_x, static current => current._x, out double x);
            bool hasY = TryResolve(_y, static current => current._y, out double y);
            if (hasX || hasY) {
                if (cursor.HasText) cursor.Chunk++;
                cursor.PendingSpace = false;
            }
            if (hasX) cursor.X = x - viewX;
            if (hasY) cursor.Baseline = y - viewY;
            if (TryResolve(_dx, static current => current._dx, out double dx)) cursor.X += dx;
            if (TryResolve(_dy, static current => current._dy, out double dy)) cursor.Baseline += dy;
            return TryResolveRotation(out double rotation) ? rotation : 0D;
        }

        internal void Advance(int count) {
            _index += Math.Max(0, count);
            _parent?.Advance(count);
        }

        private bool HasRemainingMultipleValues(IReadOnlyList<double>? values) =>
            values != null && values.Count - _index > 1;

        private bool TryResolve(
            IReadOnlyList<double>? local,
            Func<SvgTextPositioning, IReadOnlyList<double>?> selector,
            out double value) {
            if (local != null && _index < local.Count) {
                value = local[_index];
                return true;
            }
            if (_parent != null) return _parent.TryResolve(selector(_parent), selector, out value);
            value = 0D;
            return false;
        }

        private bool TryResolveRotation(out double value) {
            if (_rotate != null && _rotate.Count > 0) {
                value = _rotate[Math.Min(_index, _rotate.Count - 1)];
                return true;
            }
            if (_parent != null) return _parent.TryResolveRotation(out value);
            value = 0D;
            return false;
        }

        private static IReadOnlyList<double>? ParseLengthList(
            XElement element,
            string name,
            double percentageReference,
            ref int unsupported) {
            string? text = element.Attribute(name)?.Value;
            if (string.IsNullOrWhiteSpace(text)) return null;
            string[] tokens = text!.Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0 || tokens.Length > MaximumTextRuns) {
                unsupported++;
                return null;
            }
            var values = new List<double>(tokens.Length);
            foreach (string token in tokens) {
                if (!TryViewportLength(token, percentageReference, out double value, out _)) {
                    unsupported++;
                    return null;
                }
                values.Add(value);
            }
            return values.AsReadOnly();
        }

        private static IReadOnlyList<double>? ParseRotationList(XElement element, ref int unsupported) {
            string? text = element.Attribute("rotate")?.Value;
            if (string.IsNullOrWhiteSpace(text)) return null;
            string[] tokens = text!.Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0 || tokens.Length > MaximumTextRuns) {
                unsupported++;
                return null;
            }
            var values = new List<double>(tokens.Length);
            foreach (string token in tokens) {
                if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                    || double.IsNaN(value) || double.IsInfinity(value)) {
                    unsupported++;
                    return null;
                }
                values.Add(value);
            }
            return values.AsReadOnly();
        }
    }
}
