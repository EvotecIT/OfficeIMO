using DocumentFormat.OpenXml.Wordprocessing;
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
            BookmarkStart bookmarkStart,
            ReferenceListSwitch listSwitch,
            IReadOnlyList<string> switches,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;

            Paragraph? targetParagraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();
            if (targetParagraph == null) {
                message = "REF paragraph-number switch needs a bookmark inside a paragraph.";
                return false;
            }

            Body? body = document._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            if (body == null || !targetParagraph.Ancestors<Body>().Any()) {
                message = "REF paragraph-number switches are evaluated only for bookmarks in the main document body.";
                return false;
            }

            Numbering? numbering = document._wordprocessingDocument.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) {
                message = "REF paragraph-number switch needs numbering definitions in the document.";
                return false;
            }

            Dictionary<int, ReferenceNumberingDefinition> definitions = BuildReferenceNumberingDefinitions(numbering);
            IReadOnlyDictionary<string, ReferenceParagraphNumbering> paragraphStyleNumbering = BuildReferenceParagraphStyleNumbering(document, numbering);
            if (!TryBuildReferenceListSnapshot(body, definitions, paragraphStyleNumbering, targetParagraph, out ReferenceListNumber? targetNumber, out string? diagnostic)) {
                message = diagnostic ?? "REF paragraph-number switch could not resolve the bookmarked paragraph number.";
                return false;
            }

            ReferenceListNumber? sourceNumber = null;
            if (listSwitch == ReferenceListSwitch.RelativeContext) {
                Paragraph? sourceParagraph = candidate.AnchorElement is Paragraph paragraph
                    ? paragraph
                    : candidate.AnchorElement.Ancestors<Paragraph>().FirstOrDefault();
                if (sourceParagraph != null) {
                    TryBuildReferenceListSnapshot(body, definitions, paragraphStyleNumbering, sourceParagraph, out sourceNumber, out _);
                }
            }

            bool suppressText = switches.Any(fieldSwitch => string.Equals(fieldSwitch.Trim(), "\\t", StringComparison.OrdinalIgnoreCase));
            value = FormatReferenceListNumber(targetNumber.GetValueOrDefault(), sourceNumber, listSwitch, suppressText);

            status = WordFieldUpdateStatus.Updated;
            message = "Updated REF paragraph number from deterministic numbering definitions.";
            return true;
        }

        private static Dictionary<int, ReferenceNumberingDefinition> BuildReferenceNumberingDefinitions(Numbering numbering) {
            Dictionary<int, AbstractNum> abstracts = numbering.Elements<AbstractNum>()
                .Where(abstractNum => abstractNum.AbstractNumberId?.Value != null)
                .ToDictionary(abstractNum => abstractNum.AbstractNumberId!.Value, abstractNum => abstractNum);

            var result = new Dictionary<int, ReferenceNumberingDefinition>();
            foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>()) {
                if (instance.NumberID?.Value == null) {
                    continue;
                }

                int numberId = instance.NumberID.Value;
                int? abstractId = instance.AbstractNumId?.Val?.Value;
                if (!abstractId.HasValue || !abstracts.TryGetValue(abstractId.Value, out AbstractNum? abstractNum)) {
                    continue;
                }

                Dictionary<int, int> starts = abstractNum.Elements<Level>()
                    .Where(level => level.LevelIndex?.Value != null)
                    .ToDictionary(
                        level => level.LevelIndex!.Value,
                        level => level.StartNumberingValue?.Val?.Value ?? 1);

                foreach (LevelOverride levelOverride in instance.Elements<LevelOverride>()) {
                    if (levelOverride.LevelIndex?.Value == null) {
                        continue;
                    }

                    int? overrideStart = levelOverride.GetFirstChild<StartOverrideNumberingValue>()?.Val?.Value;
                    if (overrideStart.HasValue) {
                        starts[levelOverride.LevelIndex.Value] = overrideStart.Value;
                    }
                }

                Dictionary<int, ReferenceLevelDefinition> levels = abstractNum.Elements<Level>()
                    .Where(level => level.LevelIndex?.Value != null)
                    .ToDictionary(
                        level => level.LevelIndex!.Value,
                        level => new ReferenceLevelDefinition(
                            level.NumberingFormat?.Val?.Value,
                            level.LevelText?.Val?.Value,
                            starts.TryGetValue(level.LevelIndex!.Value, out int start) ? start : 1));

                result[numberId] = new ReferenceNumberingDefinition(levels);
            }

            return result;
        }

        private static bool TryBuildReferenceListSnapshot(
            Body body,
            IReadOnlyDictionary<int, ReferenceNumberingDefinition> definitions,
            IReadOnlyDictionary<string, ReferenceParagraphNumbering> paragraphStyleNumbering,
            Paragraph targetParagraph,
            out ReferenceListNumber? targetNumber,
            out string? diagnostic) {
            targetNumber = null;
            diagnostic = null;
            var countersByNumberId = new Dictionary<int, Dictionary<int, int>>();
            var formatsByNumberId = new Dictionary<int, Dictionary<int, NumberFormatValues?>>();
            var levelTextByNumberId = new Dictionary<int, Dictionary<int, string?>>();

            foreach (Paragraph paragraph in body.Descendants<Paragraph>()) {
                if (!TryGetParagraphNumbering(paragraph, paragraphStyleNumbering, out int numberId, out int level)) {
                    if (ReferenceEquals(paragraph, targetParagraph)) {
                        diagnostic = "REF paragraph-number switch needs a bookmark in a directly or style-numbered paragraph.";
                        return false;
                    }

                    continue;
                }

                if (!definitions.TryGetValue(numberId, out ReferenceNumberingDefinition definition) ||
                    !definition.Levels.TryGetValue(level, out ReferenceLevelDefinition levelDefinition)) {
                    if (ReferenceEquals(paragraph, targetParagraph)) {
                        diagnostic = "REF paragraph-number switch could not resolve the numbering definition for the bookmarked paragraph.";
                        return false;
                    }

                    continue;
                }

                if (!IsSupportedReferenceNumberFormat(levelDefinition.NumberFormat)) {
                    if (ReferenceEquals(paragraph, targetParagraph)) {
                        diagnostic = $"REF paragraph-number switch does not support numbering format {levelDefinition.NumberFormat?.ToString() ?? "unknown"}.";
                        return false;
                    }

                    continue;
                }

                if (!countersByNumberId.TryGetValue(numberId, out Dictionary<int, int>? counters)) {
                    counters = new Dictionary<int, int>();
                    countersByNumberId[numberId] = counters;
                    formatsByNumberId[numberId] = new Dictionary<int, NumberFormatValues?>();
                    levelTextByNumberId[numberId] = new Dictionary<int, string?>();
                }

                Dictionary<int, NumberFormatValues?> formats = formatsByNumberId[numberId];
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

                if (ReferenceEquals(paragraph, targetParagraph)) {
                    targetNumber = new ReferenceListNumber(
                        numberId,
                        level,
                        new Dictionary<int, int>(counters),
                        new Dictionary<int, NumberFormatValues?>(formats),
                        new Dictionary<int, string?>(levelTexts));
                    return true;
                }
            }

            diagnostic = "REF paragraph-number switch could not match the bookmarked paragraph in document order.";
            return false;
        }

        private static IReadOnlyDictionary<string, ReferenceParagraphNumbering> BuildReferenceParagraphStyleNumbering(WordDocument document, Numbering numbering) {
            var resolved = new Dictionary<string, ReferenceParagraphNumbering>(StringComparer.OrdinalIgnoreCase);
            Styles? styles = document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;

            if (styles != null) {
                Dictionary<string, Style> paragraphStyles = styles.Elements<Style>()
                    .Where(style => style.Type?.Value == StyleValues.Paragraph && !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                    .ToDictionary(style => style.StyleId!.Value!, style => style, StringComparer.OrdinalIgnoreCase);

                var cache = new Dictionary<string, ReferenceParagraphNumbering?>(StringComparer.OrdinalIgnoreCase);
                foreach (string styleId in paragraphStyles.Keys) {
                    if (TryResolveReferenceStyleNumbering(styleId, paragraphStyles, cache, new HashSet<string>(StringComparer.OrdinalIgnoreCase), out ReferenceParagraphNumbering styleNumbering)) {
                        resolved[styleId] = styleNumbering;
                    }
                }
            }

            AddReferenceNumberingLevelParagraphStyles(numbering, resolved);
            return resolved;
        }

        private static void AddReferenceNumberingLevelParagraphStyles(
            Numbering numbering,
            IDictionary<string, ReferenceParagraphNumbering> resolved) {
            Dictionary<int, AbstractNum> abstracts = numbering.Elements<AbstractNum>()
                .Where(abstractNum => abstractNum.AbstractNumberId?.Value != null)
                .ToDictionary(abstractNum => abstractNum.AbstractNumberId!.Value, abstractNum => abstractNum);

            foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>()) {
                if (instance.NumberID?.Value == null) {
                    continue;
                }

                int? abstractId = instance.AbstractNumId?.Val?.Value;
                if (!abstractId.HasValue || !abstracts.TryGetValue(abstractId.Value, out AbstractNum? abstractNum)) {
                    continue;
                }

                foreach (Level level in abstractNum.Elements<Level>()) {
                    int? levelIndex = level.LevelIndex?.Value;
                    string? paragraphStyleId = level.GetFirstChild<ParagraphStyleIdInLevel>()?.Val?.Value;
                    if (!levelIndex.HasValue || string.IsNullOrWhiteSpace(paragraphStyleId) || resolved.ContainsKey(paragraphStyleId!)) {
                        continue;
                    }

                    resolved[paragraphStyleId!] = new ReferenceParagraphNumbering(instance.NumberID.Value, levelIndex.Value);
                }
            }
        }

        private static bool TryResolveReferenceStyleNumbering(
            string styleId,
            IReadOnlyDictionary<string, Style> paragraphStyles,
            IDictionary<string, ReferenceParagraphNumbering?> cache,
            ISet<string> visiting,
            out ReferenceParagraphNumbering numbering) {
            numbering = default;

            if (cache.TryGetValue(styleId, out ReferenceParagraphNumbering? cached)) {
                if (cached.HasValue) {
                    numbering = cached.Value;
                    return true;
                }

                return false;
            }

            if (!paragraphStyles.TryGetValue(styleId, out Style? style) || !visiting.Add(styleId)) {
                cache[styleId] = null;
                return false;
            }

            ReferenceParagraphNumbering? inherited = null;
            string? baseStyleId = style.BasedOn?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(baseStyleId) &&
                TryResolveReferenceStyleNumbering(baseStyleId!, paragraphStyles, cache, visiting, out ReferenceParagraphNumbering baseNumbering)) {
                inherited = baseNumbering;
            }

            NumberingProperties? numberingProperties = style.GetFirstChild<StyleParagraphProperties>()?.NumberingProperties;
            int? numberId = numberingProperties?.NumberingId?.Val?.Value;
            int? level = numberingProperties?.NumberingLevelReference?.Val?.Value;

            visiting.Remove(styleId);

            if (numberId.HasValue || level.HasValue) {
                if (!numberId.HasValue && !inherited.HasValue) {
                    cache[styleId] = null;
                    return false;
                }

                numbering = new ReferenceParagraphNumbering(
                    numberId ?? inherited!.Value.NumberId,
                    level ?? inherited?.Level ?? 0);
                cache[styleId] = numbering;
                return true;
            }

            if (inherited.HasValue) {
                numbering = inherited.Value;
                cache[styleId] = numbering;
                return true;
            }

            cache[styleId] = null;
            return false;
        }

        private static bool TryGetParagraphNumbering(
            Paragraph paragraph,
            IReadOnlyDictionary<string, ReferenceParagraphNumbering> paragraphStyleNumbering,
            out int numberId,
            out int level) {
            numberId = 0;
            level = 0;
            NumberingProperties? numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
            int? directNumberId = numberingProperties?.NumberingId?.Val?.Value;
            int? directLevel = numberingProperties?.NumberingLevelReference?.Val?.Value;

            ReferenceParagraphNumbering? styleNumbering = null;
            string? paragraphStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
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
                targetNumber.Formats.TryGetValue(targetNumber.Level, out NumberFormatValues? format);
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

                number.Formats.TryGetValue(level, out NumberFormatValues? format);
                parts.Add(FormatReferenceNumber(counter, format));
            }

            if (parts.Count == 0 && number.Counters.TryGetValue(number.Level, out int current)) {
                number.Formats.TryGetValue(number.Level, out NumberFormatValues? format);
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

                number.Formats.TryGetValue(level, out NumberFormatValues? format);
                result = result.Replace("%" + (level + 1).ToString(CultureInfo.InvariantCulture), FormatReferenceNumber(counter, format));
            }

            return result;
        }

        private static string FormatReferenceNumber(int value, NumberFormatValues? format) {
            if (format == NumberFormatValues.UpperRoman) {
                return ToRoman(value).ToUpperInvariant();
            }

            if (format == NumberFormatValues.LowerRoman) {
                return ToRoman(value).ToLowerInvariant();
            }

            if (format == NumberFormatValues.UpperLetter) {
                return ToAlphabetic(value, uppercase: true);
            }

            if (format == NumberFormatValues.LowerLetter) {
                return ToAlphabetic(value, uppercase: false);
            }

            return value.ToString(CultureInfo.InvariantCulture);
        }

        private static bool IsSupportedReferenceNumberFormat(NumberFormatValues? format) {
            return format == null ||
                format == NumberFormatValues.Decimal ||
                format == NumberFormatValues.UpperRoman ||
                format == NumberFormatValues.LowerRoman ||
                format == NumberFormatValues.UpperLetter ||
                format == NumberFormatValues.LowerLetter;
        }

        private static string TrimTrailingReferenceDelimiters(string value) {
            return value.TrimEnd('.', ')', ' ', '\t');
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
            internal ReferenceLevelDefinition(NumberFormatValues? numberFormat, string? levelText, int start) {
                NumberFormat = numberFormat;
                LevelText = levelText;
                Start = start;
            }

            internal NumberFormatValues? NumberFormat { get; }

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
                IReadOnlyDictionary<int, NumberFormatValues?> formats,
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

            internal IReadOnlyDictionary<int, NumberFormatValues?> Formats { get; }

            internal IReadOnlyDictionary<int, string?> LevelTexts { get; }
        }
    }
}
