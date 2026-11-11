namespace OfficeIMO.Excel {
    internal static class ExcelHeaderNameHelper {
        internal static string[] BuildUniqueHeaders(int columnCount, Func<int, string?> headerFactory, bool normalizeHeaders) {
            ArgumentNullException.ThrowIfNull(headerFactory);

            var headers = new string[columnCount];
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var suffixes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var explicitNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int c = 0; c < columnCount; c++) {
                string normalized = NormalizeHeader(headerFactory(c), normalizeHeaders);
                if (!string.IsNullOrEmpty(normalized)) {
                    explicitNames.Add(normalized);
                }
            }

            for (int c = 0; c < columnCount; c++) {
                string normalized = NormalizeHeader(headerFactory(c), normalizeHeaders);
                bool generated = string.IsNullOrEmpty(normalized);
                string baseName = normalized;
                if (generated) {
                    baseName = $"Column{c + 1}";
                }

                bool preferBaseName = !generated || !explicitNames.Contains(baseName);
                headers[c] = MakeUnique(baseName, usedNames, suffixes, explicitNames, preferBaseName, generated);
            }

            return headers;
        }

        internal static string NormalizeHeader(string? header, bool normalizeHeaders) {
            string value = header ?? string.Empty;
            if (normalizeHeaders) {
                value = CollapseWhitespace(value);
            }

            return value;
        }

        // Collapses each whitespace run to one space and trims, as Regex.Replace(value, @"\s+", " ").Trim() does
        // (both use char.IsWhiteSpace), without building a regex for every header.
        private static string CollapseWhitespace(string value) {
            int index = 0;
            while (index < value.Length && (!char.IsWhiteSpace(value[index])
                || (value[index] == ' ' && (index + 1 == value.Length || !char.IsWhiteSpace(value[index + 1]))))) {
                index++;
            }
            if (index == value.Length) return value.Trim();

            var builder = new System.Text.StringBuilder(value.Length);
            builder.Append(value, 0, index);
            bool inWhitespace = false;
            for (; index < value.Length; index++) {
                char current = value[index];
                if (!char.IsWhiteSpace(current)) {
                    builder.Append(current);
                    inWhitespace = false;
                } else if (!inWhitespace) {
                    builder.Append(' ');
                    inWhitespace = true;
                }
            }
            return builder.ToString().Trim();
        }

        private static string MakeUnique(
            string baseName,
            HashSet<string> usedNames,
            Dictionary<string, int> suffixes,
            HashSet<string> explicitNames,
            bool preferBaseName,
            bool generated) {
            if (preferBaseName && usedNames.Add(baseName)) {
                if (!suffixes.ContainsKey(baseName)) {
                    suffixes[baseName] = 1;
                }
                return baseName;
            }

            int count = suffixes.TryGetValue(baseName, out int existingCount) ? Math.Max(existingCount + 1, 2) : 2;
            while (true) {
                string candidate = $"{baseName}_{count}";
                if ((generated && explicitNames.Contains(candidate)) || !usedNames.Add(candidate)) {
                    count++;
                    continue;
                }

                suffixes[baseName] = count;
                return candidate;
            }
        }
    }
}
