using System;
using System.Collections.Generic;

namespace OfficeIMO.Examples.Utils {
    internal static class Guard {
        internal static T NotNull<T>(T? value, string message) where T : class {
            if (value is null) {
                throw new InvalidOperationException(message);
            }

            return value;
        }

        internal static string NotNullOrWhiteSpace(string? value, string message) {
            if (string.IsNullOrWhiteSpace(value)) {
                throw new InvalidOperationException(message);
            }

            return value;
        }

        internal static T GetRequiredItem<T>(IReadOnlyList<T> values, int index, string message) {
            if (index < 0 || index >= values.Count) {
                throw new InvalidOperationException(message);
            }

            return values[index];
        }
    }
}
