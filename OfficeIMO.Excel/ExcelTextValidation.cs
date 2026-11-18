namespace OfficeIMO.Excel {
    /// <summary>Validates workbook text that every Excel file format stores as well-formed Unicode.</summary>
    internal static class ExcelTextValidation {
        /// <summary>
        /// Rejects text with a UTF-16 surrogate that is not part of a pair. XML cannot carry it, and the binary
        /// formats would store text no reader can decode, so the workbook fails late at save time or opens damaged.
        /// </summary>
        internal static void ThrowIfUnpairedSurrogate(string? value, string parameterName, string description) {
            if (value != null && FindUnpairedSurrogate(value) is int index) {
                throw new ArgumentException(
                    $"{description} contains an unpaired UTF-16 surrogate at position {index}.", parameterName);
            }
        }

        /// <summary>Returns the index of the first surrogate that is not part of a pair, or null when there is none.</summary>
        internal static int? FindUnpairedSurrogate(string value) {
            for (int index = 0; index < value.Length; index++) {
                char current = value[index];
                if (char.IsHighSurrogate(current)) {
                    if (index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) {
                        index++;
                        continue;
                    }

                    return index;
                }

                if (char.IsLowSurrogate(current)) {
                    return index;
                }
            }

            return null;
        }
    }
}
