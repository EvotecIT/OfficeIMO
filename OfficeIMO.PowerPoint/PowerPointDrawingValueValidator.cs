using System;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointDrawingValueValidator {
        internal const long MinimumDrawingCoordinate = -27273042329600L;
        internal const long MaximumDrawingCoordinate = 27273042316900L;

        internal static bool IsCoordinateInRange(long value) =>
            value >= MinimumDrawingCoordinate && value <= MaximumDrawingCoordinate;

        internal static void ValidateCoordinate(long value,
            string parameterName, string valueDescription) {
            if (!IsCoordinateInRange(value)) {
                throw new ArgumentOutOfRangeException(parameterName,
                    $"{valueDescription} must be between {MinimumDrawingCoordinate} and {MaximumDrawingCoordinate} EMUs.");
            }
        }
    }
}
