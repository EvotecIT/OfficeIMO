namespace OfficeIMO.AsciiDoc;

internal static class AsciiDocIncludeSelector {
    internal static string Apply(string source, AsciiDocElementAttributes attributes) {
        string selected = source;
        string? lines = attributes.GetNamedValue("lines");
        if (!string.IsNullOrWhiteSpace(lines)) selected = SelectLines(selected, lines!);

        string? tags = attributes.GetNamedValue("tags") ?? attributes.GetNamedValue("tag");
        if (!string.IsNullOrWhiteSpace(tags)) selected = SelectTags(selected, tags!);

        string? levelOffset = attributes.GetNamedValue("leveloffset");
        if (!string.IsNullOrWhiteSpace(levelOffset) && TryParseSignedInteger(levelOffset!, out int offset) && offset != 0) {
            selected = ApplyLevelOffset(selected, offset);
        }
        return selected;
    }

    private static string SelectLines(string source, string specification) {
        IReadOnlyList<AsciiDocSourceLine> lines = AsciiDocLineReader.Read(source);
        var selected = new HashSet<int>();
        string[] parts = specification.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < parts.Length; index++) {
            string part = parts[index].Trim();
            int range = part.IndexOf("..", StringComparison.Ordinal);
            if (range < 0) {
                if (int.TryParse(part, out int single) && single > 0) selected.Add(single);
                continue;
            }
            string startText = part.Substring(0, range).Trim();
            string endText = part.Substring(range + 2).Trim();
            int start = startText.Length == 0 ? 1 : (int.TryParse(startText, out int parsedStart) ? parsedStart : -1);
            int end = endText.Length == 0 || endText == "-1" ? lines.Count : (int.TryParse(endText, out int parsedEnd) ? parsedEnd : -1);
            if (start < 1 || end < start) continue;
            for (int line = start; line <= end && line <= lines.Count; line++) selected.Add(line);
        }

        var output = new StringBuilder();
        for (int index = 0; index < lines.Count; index++) {
            if (selected.Contains(index + 1)) output.Append(lines[index].FullText);
        }
        return output.ToString();
    }

    private static string SelectTags(string source, string specification) {
        TagFilter filter = TagFilter.Parse(specification);
        var active = new Dictionary<string, int>(StringComparer.Ordinal);
        IReadOnlyList<AsciiDocSourceLine> lines = AsciiDocLineReader.Read(source);
        var output = new StringBuilder();
        for (int index = 0; index < lines.Count; index++) {
            AsciiDocSourceLine line = lines[index];
            if (TryGetTagMarker(line.Content, out string name, out bool isStart)) {
                active.TryGetValue(name, out int depth);
                if (isStart) active[name] = depth + 1;
                else if (depth <= 1) active.Remove(name);
                else active[name] = depth - 1;
                continue;
            }
            if (filter.IsSelected(active.Keys)) output.Append(line.FullText);
        }
        return output.ToString();
    }

    private static bool TryGetTagMarker(string content, out string name, out bool isStart) {
        name = string.Empty;
        isStart = false;
        int tag = FindDirective(content, "tag::");
        int end = FindDirective(content, "end::");
        int directive;
        string marker;
        if (tag >= 0 && (end < 0 || tag < end)) {
            directive = tag;
            marker = "tag::";
            isStart = true;
        } else if (end >= 0) {
            directive = end;
            marker = "end::";
        } else return false;

        int nameStart = directive + marker.Length;
        int closing = content.IndexOf("[]", nameStart, StringComparison.Ordinal);
        if (closing <= nameStart) return false;
        int after = closing + 2;
        if (after < content.Length && !char.IsWhiteSpace(content[after])) return false;
        string candidate = content.Substring(nameStart, closing - nameStart);
        if (candidate.Any(static character => char.IsWhiteSpace(character))) return false;
        name = candidate;
        return true;
    }

    private static int FindDirective(string content, string marker) {
        int search = 0;
        while (search < content.Length) {
            int index = content.IndexOf(marker, search, StringComparison.Ordinal);
            if (index < 0) return -1;
            if (index == 0 || !IsWordCharacter(content[index - 1])) return index;
            search = index + marker.Length;
        }
        return -1;
    }

    private static bool IsWordCharacter(char value) =>
        (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z') ||
        (value >= '0' && value <= '9') || value == '_';

    private static string ApplyLevelOffset(string source, int offset) {
        IReadOnlyList<AsciiDocSourceLine> lines = AsciiDocLineReader.Read(source);
        var output = new StringBuilder(source.Length);
        for (int index = 0; index < lines.Count; index++) {
            AsciiDocSourceLine line = lines[index];
            if (AsciiDocLineClassifier.TryParseHeading(line.Content, out int markerLength, out int titleStart)) {
                int adjusted = Math.Max(1, Math.Min(6, markerLength + offset));
                output.Append(new string('=', adjusted)).Append(line.Content.Substring(titleStart - 1)).Append(line.LineEnding);
            } else {
                output.Append(line.FullText);
            }
        }
        return output.ToString();
    }

    private static bool TryParseSignedInteger(string value, out int result) {
        if (value.Length > 1 && value[0] == '+') return int.TryParse(value.Substring(1), out result);
        return int.TryParse(value, out result);
    }

    private sealed class TagFilter {
        private readonly bool _baseAll;
        private readonly bool _includeTagged;
        private readonly bool _excludeTagged;
        private readonly HashSet<string> _included;
        private readonly HashSet<string> _excluded;

        private TagFilter(
            bool baseAll,
            bool includeTagged,
            bool excludeTagged,
            HashSet<string> included,
            HashSet<string> excluded) {
            _baseAll = baseAll;
            _includeTagged = includeTagged;
            _excludeTagged = excludeTagged;
            _included = included;
            _excluded = excluded;
        }

        internal static TagFilter Parse(string specification) {
            string[] expressions = specification.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(static expression => expression.Trim())
                .Where(static expression => expression.Length > 0)
                .ToArray();
            bool baseAll = expressions.Any(static expression => string.Equals(expression, "**", StringComparison.Ordinal));
            if (!baseAll && expressions.Length > 0 && expressions[0].StartsWith("!", StringComparison.Ordinal) &&
                !string.Equals(expressions[0], "!**", StringComparison.Ordinal)) baseAll = true;
            bool includeTagged = expressions.Any(static expression => string.Equals(expression, "*", StringComparison.Ordinal));
            bool excludeTagged = expressions.Any(static expression => string.Equals(expression, "!*", StringComparison.Ordinal));
            var included = new HashSet<string>(expressions.Where(static expression =>
                    !expression.StartsWith("!", StringComparison.Ordinal) && expression != "*" && expression != "**"),
                StringComparer.Ordinal);
            var excluded = new HashSet<string>(expressions.Where(static expression =>
                    expression.StartsWith("!", StringComparison.Ordinal) && expression != "!*" && expression != "!**")
                .Select(static expression => expression.Substring(1)), StringComparer.Ordinal);
            return new TagFilter(baseAll, includeTagged, excludeTagged, included, excluded);
        }

        internal bool IsSelected(IEnumerable<string> activeNames) {
            string[] active = activeNames.ToArray();
            bool selected = _baseAll;
            if (_includeTagged && active.Length > 0) selected = true;
            if (active.Any(name => _included.Contains(name))) selected = true;
            if (_excludeTagged && active.Length > 0 &&
                (_included.Count == 0 || active.Any(name => !_included.Contains(name)))) selected = false;
            if (active.Any(name => _excluded.Contains(name))) selected = false;
            return selected;
        }
    }
}
