using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageContentVisualParser {
    private sealed partial class Parser {
        /// <summary>
        /// Recognizes a single line without losing it to zero-length Bezier corner
        /// segments. These segments occur in square CSS border paths and do not
        /// change the line's geometry, dash progression, caps, or active clip.
        /// </summary>
        private bool TryGetSingleLinePath(out OfficePoint start, out OfficePoint end) {
            start = default;
            end = default;
            if (_pathCommands.Count < 2 || _pathCommands[0].Kind != OfficePathCommandKind.MoveTo) return false;

            OfficePoint current = _pathCommands[0].Point;
            start = current;
            bool hasLine = false;
            for (int index = 1; index < _pathCommands.Count; index++) {
                OfficePathCommand command = _pathCommands[index];
                switch (command.Kind) {
                    case OfficePathCommandKind.LineTo:
                        if (command.Point == current) continue;
                        if (hasLine) return false;
                        hasLine = true;
                        end = command.Point;
                        current = command.Point;
                        break;
                    case OfficePathCommandKind.QuadraticBezierTo:
                        if (command.ControlPoint1 != current || command.Point != current) return false;
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        if (command.ControlPoint1 != current || command.ControlPoint2 != current || command.Point != current) return false;
                        break;
                    default:
                        return false;
                }
            }
            return hasLine;
        }
    }
}
