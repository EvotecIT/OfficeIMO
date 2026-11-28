using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private enum ReferenceListSwitch {
            NoContext,
            RelativeContext,
            FullContext
        }

        private static bool TryResolveReferenceListNumber(
            WordDocument document,
            MutableFieldCandidate candidate,
            OfficeXmlNode bookmarkStart,
            FieldEvaluationState state,
            ReferenceListSwitch listSwitch,
            IReadOnlyList<string> switches,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;

            OfficeXmlNode targetParagraph = bookmarkStart.AncestorOrSelf(FieldNamespace, "p");
            if (targetParagraph.IsNone) {
                message = "REF paragraph-number switch needs a bookmark inside a paragraph.";
                return false;
            }

            OfficeXmlNode body = WordFieldInventory.BodyNode(document);
            if (body.IsNone || !targetParagraph.IsInside(body)) {
                message = "REF paragraph-number switches are evaluated only for bookmarks in the main document body.";
                return false;
            }

            OfficeXmlNode numbering = ReadNumberingRoot(document);
            if (numbering.IsNone) {
                message = "REF paragraph-number switch needs numbering definitions in the document.";
                return false;
            }

            Dictionary<int, ReferenceNumberingDefinition> definitions = state.ReferenceNumberingDefinitions ??= BuildReferenceNumberingDefinitions(numbering);
            IReadOnlyDictionary<string, ReferenceParagraphNumbering> paragraphStyleNumbering = state.ReferenceStyleNumbering ??= BuildReferenceParagraphStyleNumbering(document, numbering);
            if (!TryBuildReferenceListSnapshot(body, definitions, paragraphStyleNumbering, targetParagraph, out ReferenceListNumber? targetNumber, out string? diagnostic)) {
                message = diagnostic ?? "REF paragraph-number switch could not resolve the bookmarked paragraph number.";
                return false;
            }

            ReferenceListNumber? sourceNumber = null;
            if (listSwitch == ReferenceListSwitch.RelativeContext) {
                OfficeXmlNode sourceParagraph = candidate.AnchorElement.AncestorOrSelf(FieldNamespace, "p");
                if (sourceParagraph.Exists) {
                    TryBuildReferenceListSnapshot(body, definitions, paragraphStyleNumbering, sourceParagraph, out sourceNumber, out _);
                }
            }

            bool suppressText = switches.Any(fieldSwitch => string.Equals(fieldSwitch.Trim(), "\\t", StringComparison.OrdinalIgnoreCase));
            value = FormatReferenceListNumber(targetNumber.GetValueOrDefault(), sourceNumber, listSwitch, suppressText);

            status = WordFieldUpdateStatus.Updated;
            message = "Updated REF paragraph number from deterministic numbering definitions.";
            return true;
        }

        private static OfficeXmlNode ReadNumberingRoot(WordDocument document) =>
            WordPackageView.For(document).Main?.FirstPart(WordPackageView.NumberingRelationshipType)
                ?.RootNamed(FieldNamespace, "numbering") ?? default;

        private static Dictionary<int, OfficeXmlNode> ReadNumberingNodes(OfficeXmlNode root, string name, string idAttribute) {
            const int W = FieldNamespace;
            var values = new Dictionary<int, OfficeXmlNode>();
            foreach (OfficeXmlNode node in root.Children(W, name)) {
                if (OfficeXmlViewValues.ParseInt32(node.Attribute(W, idAttribute)) is int id && !values.ContainsKey(id))
                    values.Add(id, node);
            }
            return values;
        }

        private static string? ReadNumberingValue(OfficeXmlNode parent, string name) =>
            parent.Child(FieldNamespace, name).Attribute(FieldNamespace, "val");

        private static int? ReadNumberingInt(OfficeXmlNode parent, string name) =>
            OfficeXmlViewValues.ParseInt32(ReadNumberingValue(parent, name));

        private static Dictionary<int, ReferenceNumberingDefinition> BuildReferenceNumberingDefinitions(OfficeXmlNode numbering) {
            const int W = FieldNamespace;
            Dictionary<int, OfficeXmlNode> abstracts = ReadNumberingNodes(numbering, "abstractNum", "abstractNumId");
            var result = new Dictionary<int, ReferenceNumberingDefinition>();
            foreach (OfficeXmlNode instance in numbering.Children(W, "num")) {
                if (OfficeXmlViewValues.ParseInt32(instance.Attribute(W, "numId")) is not int id ||
                    ReadNumberingInt(instance, "abstractNumId") is not int abstractId ||
                    !abstracts.TryGetValue(abstractId, out OfficeXmlNode abstractNum)) continue;

                Dictionary<int, ReferenceLevelDefinition> levels = ReadNumberingNodes(abstractNum, "lvl", "ilvl")
                    .ToDictionary(entry => entry.Key, entry => CreateReferenceLevelDefinition(entry.Value,
                        ReadNumberingInt(entry.Value, "start") ?? 1));

                foreach (OfficeXmlNode levelOverride in instance.Children(W, "lvlOverride")) {
                    if (OfficeXmlViewValues.ParseInt32(levelOverride.Attribute(W, "ilvl")) is not int levelIndex) continue;
                    OfficeXmlNode overrideLevel = levelOverride.Child(W, "lvl");
                    int start = ReadNumberingInt(levelOverride, "startOverride") ?? ReadNumberingInt(overrideLevel, "start")
                        ?? (levels.TryGetValue(levelIndex, out ReferenceLevelDefinition existingStartLevel) ? existingStartLevel.Start : 1);

                    if (overrideLevel.Exists) {
                        levels[levelIndex] = CreateReferenceLevelDefinition(overrideLevel, start);
                    } else if (levels.TryGetValue(levelIndex, out ReferenceLevelDefinition existingLevel)) {
                        levels[levelIndex] = new ReferenceLevelDefinition(existingLevel.NumberFormat, existingLevel.LevelText, start);
                    }
                }

                result[id] = new ReferenceNumberingDefinition(levels);
            }
            return result;
        }

        private static ReferenceLevelDefinition CreateReferenceLevelDefinition(OfficeXmlNode level, int start) =>
            new(ReadNumberingValue(level, "numFmt"), ReadNumberingValue(level, "lvlText"), start);

        private static bool TryBuildReferenceListSnapshot(
            OfficeXmlNode body,
            IReadOnlyDictionary<int, ReferenceNumberingDefinition> definitions,
            IReadOnlyDictionary<string, ReferenceParagraphNumbering> paragraphStyleNumbering,
            OfficeXmlNode targetParagraph,
            out ReferenceListNumber? targetNumber,
            out string? diagnostic) {
            targetNumber = null;
            diagnostic = null;
            var countersByNumberId = new Dictionary<int, Dictionary<int, int>>();
            var formatsByNumberId = new Dictionary<int, Dictionary<int, string?>>();
            var levelTextByNumberId = new Dictionary<int, Dictionary<int, string?>>();

            foreach (OfficeXmlNode paragraph in WordFieldInventory.EnumerateFieldContent(body)) {
                if (!paragraph.Is(FieldNamespace, "p")) continue;
                if (!TryGetParagraphNumbering(paragraph, paragraphStyleNumbering, out int numberId, out int level)) {
                    if (paragraph == targetParagraph) {
                        diagnostic = "REF paragraph-number switch needs a bookmark in a directly or style-numbered paragraph.";
                        return false;
                    }

                    continue;
                }

                if (!definitions.TryGetValue(numberId, out ReferenceNumberingDefinition definition) ||
                    !definition.Levels.TryGetValue(level, out ReferenceLevelDefinition levelDefinition)) {
                    if (paragraph == targetParagraph) {
                        diagnostic = "REF paragraph-number switch could not resolve the numbering definition for the bookmarked paragraph.";
                        return false;
                    }

                    continue;
                }

                if (!IsSupportedReferenceNumberFormat(levelDefinition.NumberFormat)) {
                    if (paragraph == targetParagraph) {
                        // Keep the report text returned for known and missing formats.
                        string reportFormat = levelDefinition.NumberFormat == null ? "unknown" : "NumberFormatValues { }";
                        diagnostic = $"REF paragraph-number switch does not support numbering format {reportFormat}.";
                        return false;
                    }

                    continue;
                }

                if (!countersByNumberId.TryGetValue(numberId, out Dictionary<int, int>? counters)) {
                    counters = new Dictionary<int, int>();
                    countersByNumberId[numberId] = counters;
                    formatsByNumberId[numberId] = new Dictionary<int, string?>();
                    levelTextByNumberId[numberId] = new Dictionary<int, string?>();
                }

                Dictionary<int, string?> formats = formatsByNumberId[numberId];
                Dictionary<int, string?> levelTexts = levelTextByNumberId[numberId];
                foreach (int deeperLevel in counters.Keys.Where(existingLevel => existingLevel > level).ToArray()) {
                    counters.Remove(deeperLevel);
                    formats.Remove(deeperLevel);
                    levelTexts.Remove(deeperLevel);
                }

                int currentValue = counters.TryGetValue(level, out int existingValue)
                    ? existingValue + 1
                    : levelDefinition.Start;
                counters[level] = currentValue;
                formats[level] = levelDefinition.NumberFormat;
                levelTexts[level] = levelDefinition.LevelText;

                if (paragraph == targetParagraph) {
                    targetNumber = new ReferenceListNumber(
                        numberId,
                        level,
                        new Dictionary<int, int>(counters),
                        new Dictionary<int, string?>(formats),
                        new Dictionary<int, string?>(levelTexts));
                    return true;
                }
            }

            diagnostic = "REF paragraph-number switch could not match the bookmarked paragraph in document order.";
            return false;
        }

        private static IReadOnlyDictionary<string, ReferenceParagraphNumbering> BuildReferenceParagraphStyleNumbering(
            WordDocument document, OfficeXmlNode numbering) {
            const int W = FieldNamespace;
            var resolved = new Dictionary<string, ReferenceParagraphNumbering>(StringComparer.Ordinal);
            OfficeXmlNode styles = WordPackageView.For(document).Main?.FirstPart(WordPackageView.StylesRelationshipType)
                ?.RootNamed(W, "styles") ?? default;
            if (styles.Exists) {
                var paragraphStyles = new Dictionary<string, OfficeXmlNode>(StringComparer.Ordinal);
                foreach (OfficeXmlNode style in styles.Children(W, "style")) {
                    if (style.Attribute(W, "type") == "paragraph" && style.Attribute(W, "styleId") is { } id &&
                        !string.IsNullOrWhiteSpace(id) && !paragraphStyles.ContainsKey(id)) paragraphStyles.Add(id, style);
                }

                var cache = new Dictionary<string, ReferenceParagraphNumbering?>(StringComparer.Ordinal);
                foreach (string styleId in paragraphStyles.Keys) {
                    if (ResolveReferenceStyleNumbering(styleId, paragraphStyles, cache) is { } styleNumbering)
                        resolved[styleId] = styleNumbering;
                }
            }

            AddReferenceNumberingLevelParagraphStyles(numbering, resolved);
            return resolved;
        }

        private static void AddReferenceNumberingLevelParagraphStyles(OfficeXmlNode numbering,
            IDictionary<string, ReferenceParagraphNumbering> resolved) {
            const int W = FieldNamespace;
            Dictionary<int, OfficeXmlNode> abstracts = ReadNumberingNodes(numbering, "abstractNum", "abstractNumId");
            foreach (OfficeXmlNode instance in numbering.Children(W, "num")) {
                if (OfficeXmlViewValues.ParseInt32(instance.Attribute(W, "numId")) is not int id ||
                    ReadNumberingInt(instance, "abstractNumId") is not int abstractId ||
                    !abstracts.TryGetValue(abstractId, out OfficeXmlNode abstractNum)) continue;

                foreach (OfficeXmlNode level in abstractNum.Children(W, "lvl")) {
                    string? styleId = ReadNumberingValue(level, "pStyle");
                    if (OfficeXmlViewValues.ParseInt32(level.Attribute(W, "ilvl")) is int index &&
                        !string.IsNullOrWhiteSpace(styleId) && !resolved.ContainsKey(styleId!))
                        resolved[styleId!] = new ReferenceParagraphNumbering(id, index);
                }
            }
        }

        private static ReferenceParagraphNumbering? ResolveReferenceStyleNumbering(string styleId,
            IReadOnlyDictionary<string, OfficeXmlNode> styles, IDictionary<string, ReferenceParagraphNumbering?> cache) {
            const int W = FieldNamespace;
            if (cache.TryGetValue(styleId, out ReferenceParagraphNumbering? value)) return value;

            var pending = new List<KeyValuePair<string, OfficeXmlNode>>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            string? current = styleId;
            while (!string.IsNullOrWhiteSpace(current)) {
                if (cache.TryGetValue(current!, out value)) break;
                if (!styles.TryGetValue(current!, out OfficeXmlNode style) || !seen.Add(current!)) break;
                pending.Add(new KeyValuePair<string, OfficeXmlNode>(current!, style));
                current = ReadNumberingValue(style, "basedOn");
            }

            for (int i = pending.Count - 1; i >= 0; i--) {
                var entry = pending[i];
                OfficeXmlNode numbering = entry.Value.Child(W, "pPr").Child(W, "numPr");
                int? id = ReadNumberingInt(numbering, "numId");
                int? level = ReadNumberingInt(numbering, "ilvl");
                if ((id ?? value?.NumberId) is int effectiveId) {
                    value = new ReferenceParagraphNumbering(effectiveId, level ?? value?.Level ?? 0);
                }
                cache[entry.Key] = value;
            }
            return value;
        }

        private static bool TryGetParagraphNumbering(
            OfficeXmlNode paragraph,
            IReadOnlyDictionary<string, ReferenceParagraphNumbering> paragraphStyleNumbering,
            out int numberId,
            out int level) {
            numberId = 0;
            level = 0;
            const int W = FieldNamespace;
            OfficeXmlNode properties = paragraph.Child(W, "pPr");
            OfficeXmlNode numberingProperties = properties.Child(W, "numPr");
            int? directNumberId = ReadNumberingInt(numberingProperties, "numId");
            int? directLevel = ReadNumberingInt(numberingProperties, "ilvl");

            ReferenceParagraphNumbering? styleNumbering = null;
            string? paragraphStyleId = ReadNumberingValue(properties, "pStyle");
            if (!string.IsNullOrWhiteSpace(paragraphStyleId) &&
                paragraphStyleNumbering.TryGetValue(paragraphStyleId!, out ReferenceParagraphNumbering resolvedStyleNumbering)) {
                styleNumbering = resolvedStyleNumbering;
            }

            if (directNumberId.HasValue) {
                numberId = directNumberId.Value;
                level = directLevel ?? styleNumbering?.Level ?? 0;
                return level >= 0 && level <= 8;
            }

            if (directLevel.HasValue && styleNumbering.HasValue) {
                numberId = styleNumbering.Value.NumberId;
                level = directLevel.Value;
                return level >= 0 && level <= 8;
            }

            if (styleNumbering.HasValue) {
                numberId = styleNumbering.Value.NumberId;
                level = styleNumbering.Value.Level;
                return level >= 0 && level <= 8;
            }

            return false;
        }

        private static string FormatReferenceListNumber(
            ReferenceListNumber targetNumber,
            ReferenceListNumber? sourceNumber,
            ReferenceListSwitch listSwitch,
            bool suppressText) {
            if (listSwitch == ReferenceListSwitch.NoContext) {
                int currentValue = targetNumber.Counters[targetNumber.Level];
                targetNumber.Formats.TryGetValue(targetNumber.Level, out string? format);
                return FormatReferenceNumber(currentValue, format);
            }

            int firstLevel = 0;
            if (listSwitch == ReferenceListSwitch.RelativeContext &&
                sourceNumber.HasValue &&
                sourceNumber.Value.NumberId == targetNumber.NumberId) {
                firstLevel = CountSharedLeadingLevels(targetNumber, sourceNumber.Value);
                if (firstLevel > targetNumber.Level) {
                    firstLevel = targetNumber.Level;
                }
            }

            return BuildReferenceContextNumber(targetNumber, firstLevel, suppressText);
        }

        private static int CountSharedLeadingLevels(ReferenceListNumber targetNumber, ReferenceListNumber sourceNumber) {
            int shared = 0;
            int maxLevel = Math.Min(targetNumber.Level, sourceNumber.Level);
            for (int level = 0; level <= maxLevel; level++) {
                if (!targetNumber.Counters.TryGetValue(level, out int targetValue) ||
                    !sourceNumber.Counters.TryGetValue(level, out int sourceValue) ||
                    targetValue != sourceValue) {
                    break;
                }

                shared++;
            }

            return shared;
        }

        private static string BuildReferenceContextNumber(ReferenceListNumber number, int firstLevel, bool suppressText) {
            number.LevelTexts.TryGetValue(number.Level, out string? pattern);
            if (!suppressText && !string.IsNullOrWhiteSpace(pattern) && firstLevel == 0) {
                string expanded = ExpandReferenceLevelText(pattern!, number);
                return TrimTrailingReferenceDelimiters(expanded);
            }

            var parts = new List<string>();
            for (int level = Math.Max(0, firstLevel); level <= number.Level; level++) {
                if (!number.Counters.TryGetValue(level, out int counter)) {
                    continue;
                }

                number.Formats.TryGetValue(level, out string? format);
                parts.Add(FormatReferenceNumber(counter, format));
            }

            if (parts.Count == 0 && number.Counters.TryGetValue(number.Level, out int current)) {
                number.Formats.TryGetValue(number.Level, out string? format);
                return FormatReferenceNumber(current, format);
            }

            return string.Join(".", parts);
        }

        private static string ExpandReferenceLevelText(string pattern, ReferenceListNumber number) {
            string result = pattern;
            for (int level = 0; level <= number.Level; level++) {
                if (!number.Counters.TryGetValue(level, out int counter)) {
                    continue;
                }

                number.Formats.TryGetValue(level, out string? format);
                result = result.Replace("%" + (level + 1).ToString(CultureInfo.InvariantCulture), FormatReferenceNumber(counter, format));
            }

            return result;
        }

        private static string FormatReferenceNumber(int value, string? format) {
            if (format == "upperRoman") {
                return ToRoman(value).ToUpperInvariant();
            }

            if (format == "lowerRoman") {
                return ToRoman(value).ToLowerInvariant();
            }

            if (format == "upperLetter") {
                return ToAlphabetic(value, uppercase: true);
            }

            if (format == "lowerLetter") {
                return ToAlphabetic(value, uppercase: false);
            }

            return value.ToString(CultureInfo.InvariantCulture);
        }

        private static bool IsSupportedReferenceNumberFormat(string? format) {
            return format == null ||
                format == "decimal" ||
                format == "upperRoman" ||
                format == "lowerRoman" ||
                format == "upperLetter" ||
                format == "lowerLetter";
        }

        private static string TrimTrailingReferenceDelimiters(string value) {
            return value.TrimEnd('.', ' ', '\t');
        }

        private static string ReferenceListSwitchToFieldCode(ReferenceListSwitch listSwitch) {
            switch (listSwitch) {
                case ReferenceListSwitch.NoContext:
                    return "\\n";
                case ReferenceListSwitch.RelativeContext:
                    return "\\r";
                case ReferenceListSwitch.FullContext:
                    return "\\w";
                default:
                    return string.Empty;
            }
        }

        private readonly struct ReferenceNumberingDefinition {
            internal ReferenceNumberingDefinition(IReadOnlyDictionary<int, ReferenceLevelDefinition> levels) {
                Levels = levels;
            }

            internal IReadOnlyDictionary<int, ReferenceLevelDefinition> Levels { get; }
        }

        private readonly struct ReferenceLevelDefinition {
            internal ReferenceLevelDefinition(string? numberFormat, string? levelText, int start) {
                NumberFormat = numberFormat;
                LevelText = levelText;
                Start = start;
            }

            internal string? NumberFormat { get; }

            internal string? LevelText { get; }

            internal int Start { get; }
        }

        private readonly struct ReferenceParagraphNumbering {
            internal ReferenceParagraphNumbering(int numberId, int level) {
                NumberId = numberId;
                Level = level;
            }

            internal int NumberId { get; }

            internal int Level { get; }
        }

        private readonly struct ReferenceListNumber {
            internal ReferenceListNumber(
                int numberId,
                int level,
                IReadOnlyDictionary<int, int> counters,
                IReadOnlyDictionary<int, string?> formats,
                IReadOnlyDictionary<int, string?> levelTexts) {
                NumberId = numberId;
                Level = level;
                Counters = counters;
                Formats = formats;
                LevelTexts = levelTexts;
            }

            internal int NumberId { get; }

            internal int Level { get; }

            internal IReadOnlyDictionary<int, int> Counters { get; }

            internal IReadOnlyDictionary<int, string?> Formats { get; }

            internal IReadOnlyDictionary<int, string?> LevelTexts { get; }
        }
    }
}
