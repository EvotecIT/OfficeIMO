using System.IO;

namespace OfficeIMO.Security;

internal static class SecurityLimits {
    internal static void EnsureBufferWithinLimit(byte[] value, long maximumBytes, string parameterName) {
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes), "The byte limit must be positive.");
        if (value.LongLength > maximumBytes) {
            throw new ArgumentException(
                $"The supplied value is {value.LongLength} bytes and exceeds the configured limit of {maximumBytes} bytes.",
                parameterName);
        }
    }

    internal static void EnsureCountWithinLimit(int value, int maximum, string limitName) {
        if (maximum <= 0) throw new ArgumentOutOfRangeException(limitName, "The count limit must be positive.");
        if (value > maximum) {
            throw new InvalidDataException($"The decoded object contains {value} entries and exceeds the configured {limitName} of {maximum}.");
        }
    }
}
