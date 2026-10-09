using OfficeIMO.IWork;
using OpenXmlNumberingId = DocumentFormat.OpenXml.Wordprocessing.NumberingId;
using OpenXmlNumberingLevelReference = DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference;
using OpenXmlNumberingProperties = DocumentFormat.OpenXml.Wordprocessing.NumberingProperties;
using OpenXmlParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    private sealed class IWorkNativeListCatalog {
        private readonly WordDocument _document;
        private readonly HashSet<int> _observedLevels = new();
        private WordList? _current;

        internal IWorkNativeListCatalog(WordDocument document) {
            _document = document;
        }

        internal IWorkNativeListCatalog CreateTextScope() => new(_document);

        internal void Apply(WordParagraph paragraph, int level, string? label, string? fontName,
            IWorkListMarkerKind markerKind, bool startsNewList, int? qualifiedNumber,
            int? qualifiedTextOffset) {
            if (startsNewList || _current == null) {
                _current = WordList.AddCustomList(_document);
                _observedLevels.Clear();
            }
            WordList list = _current;
            WordListLevelKind levelKind = markerKind == IWorkListMarkerKind.Number
                ? Classify(label) : WordListLevelKind.Bullet;
            bool targetLevelExists = list.Numbering.Levels.Count > level;
            while (list.Numbering.Levels.Count <= level) {
                list.Numbering.AddLevel(new WordListLevel(levelKind));
            }
            bool firstObservation = _observedLevels.Add(level);
            if (firstObservation && targetLevelExists) {
                ReplacePlaceholderLevel(list, level, levelKind);
            }
            if (firstObservation && levelKind != WordListLevelKind.Bullet) {
                if (qualifiedNumber.HasValue) list.Numbering.Levels[level].SetStartNumberingValue(qualifiedNumber.Value);
                else if (TryParseStart(label, levelKind, out int start) && start != 1)
                    list.Numbering.Levels[level].SetStartNumberingValue(start);
            }
            if (firstObservation && levelKind == WordListLevelKind.Bullet
                && !string.IsNullOrEmpty(label)) {
                list.Numbering.Levels[level].LevelText = label!;
            }
            if (firstObservation && levelKind != WordListLevelKind.Bullet
                && IsParenthesized(label)) {
                list.Numbering.Levels[level].LevelText = "(%"
                    + (level + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ")";
            }
            if (firstObservation && fontName != null) list.FontName = fontName;
            if (qualifiedTextOffset.HasValue) {
                WordListLevel numberingLevel = list.Numbering.Levels[level];
                numberingLevel.IndentationLeft = qualifiedTextOffset.Value;
                numberingLevel.IndentationHanging = qualifiedTextOffset.Value;
                numberingLevel.LevelSuffix = WordListLevelSuffix.Tab;
            }
            OpenXmlParagraphProperties properties = paragraph._paragraph.ParagraphProperties
                ?? paragraph._paragraph.PrependChild(new OpenXmlParagraphProperties());
            properties.NumberingProperties = new OpenXmlNumberingProperties(
                new OpenXmlNumberingLevelReference { Val = level },
                new OpenXmlNumberingId { Val = list.NumberId });
        }

        private static void ReplacePlaceholderLevel(WordList list, int level,
            WordListLevelKind levelKind) {
            WordListLevel current = list.Numbering.Levels[level];
            var replacement = new WordListLevel(levelKind);
            replacement.OpenXmlElement.LevelIndex = level;
            replacement.LevelText = replacement.LevelText.Replace(
                "%CurrentLevel", "%" + (level + 1).ToString(
                    System.Globalization.CultureInfo.InvariantCulture));
            current.OpenXmlElement.InsertAfterSelf(replacement.OpenXmlElement);
            current.Remove();
        }

        private static WordListLevelKind Classify(string? label) {
            if (string.IsNullOrWhiteSpace(label)) return WordListLevelKind.Bullet;
            string marker = label!.Trim();
            bool bracket = marker.EndsWith(")", StringComparison.Ordinal);
            bool dot = marker.EndsWith(".", StringComparison.Ordinal);
            string token = MarkerToken(marker);
            if (token.All(char.IsDigit)) {
                return bracket ? WordListLevelKind.DecimalBracket
                    : dot ? WordListLevelKind.DecimalDot : WordListLevelKind.Decimal;
            }
            bool roman = token.Length > 0
                && token.All(character => "ivxlcdmIVXLCDM".IndexOf(character) >= 0)
                && (token.Length > 1 || "ivxIVX".IndexOf(token[0]) >= 0);
            if (roman) {
                bool upper = token.All(char.IsUpper);
                return upper
                    ? bracket ? WordListLevelKind.UpperRomanBracket
                        : dot ? WordListLevelKind.UpperRomanDot : WordListLevelKind.UpperRoman
                    : bracket ? WordListLevelKind.LowerRomanBracket
                        : dot ? WordListLevelKind.LowerRomanDot : WordListLevelKind.LowerRoman;
            }
            if (token.Length > 0 && token.All(char.IsLetter)) {
                bool upper = token.All(char.IsUpper);
                return upper
                    ? bracket ? WordListLevelKind.UpperLetterBracket
                        : dot ? WordListLevelKind.UpperLetterDot : WordListLevelKind.UpperLetter
                    : bracket ? WordListLevelKind.LowerLetterBracket
                        : dot ? WordListLevelKind.LowerLetterDot : WordListLevelKind.LowerLetter;
            }
            return WordListLevelKind.Bullet;
        }

        private static bool TryParseStart(string? label, WordListLevelKind kind, out int start) {
            start = 1;
            if (string.IsNullOrWhiteSpace(label)) return false;
            string token = MarkerToken(label!.Trim());
            switch (kind) {
                case WordListLevelKind.Decimal:
                case WordListLevelKind.DecimalDot:
                case WordListLevelKind.DecimalBracket:
                    return int.TryParse(token, System.Globalization.NumberStyles.None,
                        System.Globalization.CultureInfo.InvariantCulture, out start) && start > 0;
                case WordListLevelKind.UpperLetter:
                case WordListLevelKind.UpperLetterDot:
                case WordListLevelKind.UpperLetterBracket:
                case WordListLevelKind.LowerLetter:
                case WordListLevelKind.LowerLetterDot:
                case WordListLevelKind.LowerLetterBracket:
                    start = 0;
                    foreach (char character in token) {
                        int digit = char.ToUpperInvariant(character) - 'A' + 1;
                        if (digit < 1 || digit > 26 || start > (int.MaxValue - digit) / 26) return false;
                        start = start * 26 + digit;
                    }
                    return start > 0;
                case WordListLevelKind.UpperRoman:
                case WordListLevelKind.UpperRomanDot:
                case WordListLevelKind.UpperRomanBracket:
                case WordListLevelKind.LowerRoman:
                case WordListLevelKind.LowerRomanDot:
                case WordListLevelKind.LowerRomanBracket:
                    return TryParseRoman(token, out start);
                default:
                    return false;
            }
        }

        private static bool TryParseRoman(string token, out int value) {
            value = 0;
            int previous = 0;
            for (int index = token.Length - 1; index >= 0; index--) {
                int current = char.ToUpperInvariant(token[index]) switch {
                    'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50,
                    'C' => 100, 'D' => 500, 'M' => 1000, _ => 0
                };
                if (current == 0) return false;
                int delta = current < previous ? -current : current;
                if (delta > 0 && value > int.MaxValue - delta
                    || delta < 0 && value < int.MinValue - delta) return false;
                value += delta;
                if (current > previous) previous = current;
            }
            if (value <= 0) return false;
            bool upper = token.All(character => character is >= 'A' and <= 'Z');
            string canonical = FormatRoman(value);
            return string.Equals(token, upper ? canonical : canonical.ToLowerInvariant(),
                StringComparison.Ordinal);
        }

        private static string FormatRoman(int value) {
            var builder = new System.Text.StringBuilder();
            foreach ((int Number, string Token) part in new[] {
                         (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                         (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                         (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
                     }) {
                while (value >= part.Number) {
                    builder.Append(part.Token);
                    value -= part.Number;
                }
            }
            return builder.ToString();
        }

        internal static bool CanPreserveStart(string? label) {
            if (!string.IsNullOrWhiteSpace(label)) {
                string token = MarkerToken(label!.Trim());
                if (token.All(char.IsLetter)
                    && token.Any(char.IsUpper) && token.Any(char.IsLower)) return false;
            }
            WordListLevelKind kind = Classify(label);
            return kind == WordListLevelKind.Bullet || TryParseStart(label, kind, out _);
        }

        private static bool IsParenthesized(string? marker) {
            if (string.IsNullOrWhiteSpace(marker)) return false;
            string trimmed = marker!.Trim();
            return trimmed.Length >= 2 && trimmed[0] == '('
                && trimmed[trimmed.Length - 1] == ')';
        }

        private static string MarkerToken(string marker) {
            string trimmed = marker.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '('
                && trimmed[trimmed.Length - 1] == ')') {
                return trimmed.Substring(1, trimmed.Length - 2).Trim();
            }
            return trimmed.TrimEnd('.', ')', ' ');
        }
    }

}
