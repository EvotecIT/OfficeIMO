namespace OfficeIMO.Excel {
    internal static class ExcelHeaderNameHelper {
        internal static string[] BuildUniqueHeaders(int columnCount, Func<int, string?> headerFactory, bool normalizeHeaders) {
            if (headerFactory == null) throw new ArgumentNullException(nameof(headerFactory));

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
                value = System.Text.RegularExpressions.Regex.Replace(value, "\\s+", " ").Trim();
            }

            return value;
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
