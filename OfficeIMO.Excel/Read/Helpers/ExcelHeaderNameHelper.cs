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
            return normalizeHeaders ? CollapseWhiteSpace(value) : value;
        }

        /// <summary>
        /// Trims the header and turns each inner run of white space into one space. A regular
        /// expression would cost a cold read several milliseconds to build.
        /// </summary>
        private static string CollapseWhiteSpace(string value) {
            int start = 0;
            int end = value.Length;
            while (start < end && char.IsWhiteSpace(value[start])) start++;
            while (end > start && char.IsWhiteSpace(value[end - 1])) end--;

            bool collapsed = true;
            for (int index = start; index < end && collapsed; index++) {
                // After trimming, white space is always followed by another character.
                collapsed = !char.IsWhiteSpace(value[index]) || value[index] == ' ' && !char.IsWhiteSpace(value[index + 1]);
            }
            if (collapsed) {
                return start == 0 && end == value.Length ? value : value.Substring(start, end - start);
            }

            var builder = new System.Text.StringBuilder(end - start);
            bool pendingSpace = false;
            for (int index = start; index < end; index++) {
                char character = value[index];
                if (char.IsWhiteSpace(character)) {
                    pendingSpace = true;
                    continue;
                }
                if (pendingSpace) {
                    builder.Append(' ');
                    pendingSpace = false;
                }
                builder.Append(character);
            }
            return builder.ToString();
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
