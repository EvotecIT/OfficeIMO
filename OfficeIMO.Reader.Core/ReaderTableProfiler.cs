using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Reader;

/// <summary>
/// Builds reusable column profiles for <see cref="ReaderTable"/> instances.
/// </summary>
public static class ReaderTableProfiler {
    /// <summary>
    /// Creates inferred column profiles aligned to the supplied column names.
    /// </summary>
    /// <param name="columns">Column names in table order.</param>
    /// <param name="rows">Body rows aligned to <paramref name="columns"/>.</param>
    /// <returns>Profiles describing empty, numeric, text, or mixed body values for each column.</returns>
    public static IReadOnlyList<ReaderTableColumnProfile> CreateProfiles(IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyList<string>> rows) {
        if (columns == null) throw new ArgumentNullException(nameof(columns));

        if (columns.Count == 0) {
            return Array.Empty<ReaderTableColumnProfile>();
        }

        rows ??= Array.Empty<IReadOnlyList<string>>();

        var profiles = new ReaderTableColumnProfile[columns.Count];
        for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++) {
            int nonEmptyCount = 0;
            int numericCount = 0;

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                var row = rows[rowIndex];
                if (row == null || columnIndex >= row.Count) {
                    continue;
                }

                var value = row[columnIndex];
                if (string.IsNullOrWhiteSpace(value)) {
                    continue;
                }

                nonEmptyCount++;
                if (IsNumeric(value)) {
                    numericCount++;
                }
            }

            profiles[columnIndex] = new ReaderTableColumnProfile {
                Index = columnIndex,
                Name = columns[columnIndex] ?? string.Empty,
                Kind = GetColumnKind(nonEmptyCount, numericCount),
                NonEmptyCellCount = nonEmptyCount,
                NumericCellCount = numericCount,
                Confidence = GetConfidence(nonEmptyCount, numericCount)
            };
        }

        return profiles;
    }

    private static ReaderTableColumnKind GetColumnKind(int nonEmptyCount, int numericCount) {
        if (nonEmptyCount == 0) {
            return ReaderTableColumnKind.Empty;
        }

        if (numericCount == nonEmptyCount) {
            return ReaderTableColumnKind.Numeric;
        }

        if (numericCount == 0) {
            return ReaderTableColumnKind.Text;
        }

        return ReaderTableColumnKind.Mixed;
    }

    private static double GetConfidence(int nonEmptyCount, int numericCount) {
        if (nonEmptyCount == 0) {
            return 0d;
        }

        if (numericCount == 0 || numericCount == nonEmptyCount) {
            return 1d;
        }

        var numericShare = (double)numericCount / nonEmptyCount;
        return Math.Max(numericShare, 1d - numericShare);
    }

    private static bool IsNumeric(string value) {
        return decimal.TryParse(
            value.Trim(),
            NumberStyles.Number | NumberStyles.AllowCurrencySymbol,
            CultureInfo.InvariantCulture,
            out _);
    }
}
