using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Visio {
/// <summary>
    /// Line-oriented difference between two inspection snapshots.
    /// </summary>
    public sealed class VisioInspectionDiff {
        private VisioInspectionDiff(IReadOnlyList<VisioInspectionDifference> differences) {
            Differences = differences;
        }

        /// <summary>Snapshot differences.</summary>
        public IReadOnlyList<VisioInspectionDifference> Differences { get; }

        /// <summary>Whether any snapshot line changed.</summary>
        public bool HasDifferences => Differences.Count > 0;

        /// <summary>Compares two snapshots.</summary>
        public static VisioInspectionDiff Compare(VisioInspectionSnapshot expected, VisioInspectionSnapshot actual) {
            if (expected == null) {
                throw new ArgumentNullException(nameof(expected));
            }

            if (actual == null) {
                throw new ArgumentNullException(nameof(actual));
            }

            SortedDictionary<string, string> expectedLines = ToLineMap(expected.ToText());
            SortedDictionary<string, string> actualLines = ToLineMap(actual.ToText());
            SortedSet<string> keys = new(expectedLines.Keys, StringComparer.Ordinal);
            keys.UnionWith(actualLines.Keys);

            List<VisioInspectionDifference> differences = new();
            foreach (string key in keys) {
                bool hasExpected = expectedLines.TryGetValue(key, out string? expectedValue);
                bool hasActual = actualLines.TryGetValue(key, out string? actualValue);
                if (!hasExpected && hasActual) {
                    differences.Add(new VisioInspectionDifference(VisioInspectionDifferenceKind.Added, key, null, actualValue));
                } else if (hasExpected && !hasActual) {
                    differences.Add(new VisioInspectionDifference(VisioInspectionDifferenceKind.Removed, key, expectedValue, null));
                } else if (!string.Equals(expectedValue, actualValue, StringComparison.Ordinal)) {
                    differences.Add(new VisioInspectionDifference(VisioInspectionDifferenceKind.Changed, key, expectedValue, actualValue));
                }
            }

            return new VisioInspectionDiff(differences.AsReadOnly());
        }

        /// <summary>Writes a stable text representation of the diff.</summary>
        public string ToText() {
            StringBuilder builder = new();
            foreach (VisioInspectionDifference difference in Differences) {
                builder.Append(difference.Kind);
                builder.Append(' ');
                builder.Append(difference.Path);
                builder.Append(" expected=");
                builder.Append(VisioInspectionSnapshot.FormatLineValue(difference.Expected));
                builder.Append(" actual=");
                builder.Append(VisioInspectionSnapshot.FormatLineValue(difference.Actual));
                builder.AppendLine();
            }

            return builder.ToString();
        }

        /// <inheritdoc />
        public override string ToString() {
            return ToText();
        }

        private static SortedDictionary<string, string> ToLineMap(string text) {
            SortedDictionary<string, string> map = new(StringComparer.Ordinal);
            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            foreach (string line in lines) {
                if (line.Length == 0) {
                    continue;
                }

                int separator = FindSeparator(line);
                string key = separator >= 0 ? line.Substring(0, separator) : line;
                string value = separator >= 0
                    ? VisioInspectionSnapshot.UnescapeValue(line.Substring(separator + 1))
                    : string.Empty;
                map[key] = value;
            }

            return map;
        }

        private static int FindSeparator(string line) {
            bool escaped = false;
            for (int i = 0; i < line.Length; i++) {
                char c = line[i];
                if (escaped) {
                    escaped = false;
                    continue;
                }

                if (c == '\\') {
                    escaped = true;
                    continue;
                }

                if (c == '=') {
                    return i;
                }
            }

            return -1;
        }
    }

/// <summary>
    /// Kind of inspection snapshot difference.
    /// </summary>
    public enum VisioInspectionDifferenceKind {
        /// <summary>A path exists only in the actual snapshot.</summary>
        Added,

        /// <summary>A path exists only in the expected snapshot.</summary>
        Removed,

        /// <summary>A path exists in both snapshots with a different value.</summary>
        Changed
    }

/// <summary>
    /// One inspection snapshot difference.
    /// </summary>
    public sealed class VisioInspectionDifference {
        internal VisioInspectionDifference(VisioInspectionDifferenceKind kind, string path, string? expected, string? actual) {
            Kind = kind;
            Path = path;
            Expected = expected;
            Actual = actual;
        }

        /// <summary>Difference kind.</summary>
        public VisioInspectionDifferenceKind Kind { get; }

        /// <summary>Stable snapshot path that changed.</summary>
        public string Path { get; }

        /// <summary>Expected value.</summary>
        public string? Expected { get; }

        /// <summary>Actual value.</summary>
        public string? Actual { get; }
    }
}
