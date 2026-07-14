using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private int? _orderedListNumberId;

        private void ProcessList(IElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, WordTableCell? cell, TextFormatting formatting, WordHeaderFooter? headerFooter) {
            bool ordered = element.TagName.Equals("ol", System.StringComparison.OrdinalIgnoreCase);
            int parentDepth = listStack.Count;

            WordList CreateOrderedList(bool allowContinue = true) {
                WordList listInstance;
                if (allowContinue && options.ContinueNumbering && _orderedListNumberId.HasValue && parentDepth == 0 && cell == null && headerFooter == null) {
                    listInstance = new WordList(doc);
                    listInstance.NumberId = _orderedListNumberId.Value;
                } else {
                    // Use standard numbered list style for ordered lists in all contexts
                    listInstance = cell != null ? cell.AddList(WordListStyle.Numbered)
                        : headerFooter != null ? headerFooter.AddList(WordListStyle.Numbered)
                        : doc.AddListNumbered();
                    if (allowContinue && options.ContinueNumbering && parentDepth == 0 && cell == null && headerFooter == null) {
                        _orderedListNumberId = listInstance.NumberId;
                    }
                }
                return listInstance;
            }

            WordList CreateBulletedList() {
                return cell != null ? cell.AddList(WordListStyle.Bulleted)
                    : headerFooter != null ? headerFooter.AddList(WordListStyle.Bulleted)
                    : doc.AddListBulleted();
            }

            string? listStyleType = GetListStyleType(element);
            string? typeAttr = element.GetAttribute("type");

            WordList list = ordered ? CreateOrderedList() : CreateBulletedList();
            ApplyListStyle(list, parentDepth, ordered, listStyleType, typeAttr);
            ApplyListIndentMetadata(list, parentDepth, element);

            if (ordered) {
                int? startValue = null;
                var start = element.GetAttribute("start");
                if (!string.IsNullOrEmpty(start) && int.TryParse(start, out int startVal)) {
                    startValue = startVal;
                } else if (element.HasAttribute("reversed")) {
                    var itemCount = element.Children.OfType<IHtmlListItemElement>().Count();
                    if (itemCount > 0) {
                        startValue = itemCount;
                    }
                }
                if (startValue.HasValue) {
                    list.SetStartNumberingValue(startValue.Value, parentDepth);
                }
            }

            listStack.Push(list);

            int itemIndex = 0;
            WordParagraph? insertionAnchor = null;
            foreach (var li in element.Children.OfType<IHtmlListItemElement>()) {
                if (ordered) {
                    if (TryGetListItemValue(li, out int liValue)) {
                        if (itemIndex == 0) {
                            list.SetStartNumberingValue(liValue, parentDepth);
                        } else {
                            list = CreateOrderedList(allowContinue: false);
                            ApplyListStyle(list, parentDepth, ordered, listStyleType, typeAttr);
                            ApplyListIndentMetadata(list, parentDepth, element);
                            list.SetStartNumberingValue(liValue, parentDepth);
                            listStack.Pop();
                            listStack.Push(list);
                        }
                    }
                }

                insertionAnchor = ProcessListItem(li, doc, section, options, listStack, formatting, cell, headerFooter, insertionAnchor);
                itemIndex++;
            }

            listStack.Pop();
        }

        private static string? GetListStyleType(IElement element) {
            var style = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style)) {
                return null;
            }
            foreach (var part in (style ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                var pieces = part.Split(new[] { ':' }, 2);
                if (pieces.Length != 2) {
                    continue;
                }
                var name = pieces[0].Trim().ToLowerInvariant();
                var value = pieces[1].Trim();
                if (name == "list-style-type") {
                    return NormalizeListStyleToken(value);
                }
                if (name == "list-style") {
                    var token = ExtractListStyleToken(value);
                    if (!string.IsNullOrEmpty(token)) {
                        return token;
                    }
                }
            }
            return null;
        }

        private static string? ExtractListStyleToken(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }
            var quoted = ExtractQuotedListStyleToken(value);
            if (quoted != null) {
                return quoted;
            }
            var tokens = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var raw in tokens) {
                var token = NormalizeListStyleToken(raw);
                if (token != null) {
                    return token;
                }
            }
            return null;
        }

        private static string? ExtractQuotedListStyleToken(string value) {
            char quote = '\0';
            int start = -1;
            for (int i = 0; i < value.Length; i++) {
                if (value[i] == '\'' || value[i] == '"') {
                    quote = value[i];
                    start = i + 1;
                    break;
                }
            }
            if (start < 0) {
                return null;
            }

            for (int i = start; i < value.Length; i++) {
                if (value[i] == quote) {
                    return NormalizeListStyleToken(value.Substring(start, i - start));
                }
            }

            return null;
        }

        private static string? NormalizeListStyleToken(string? value) {
            if (value == null) {
                return null;
            }
            var trimmed = value.Trim();
            if (trimmed.Length == 0) {
                return null;
            }
            var importantIndex = trimmed.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
            if (importantIndex >= 0) {
                trimmed = trimmed.Substring(0, importantIndex).Trim();
            }
            var token = trimmed.TrimEnd(',');
            var quoted = false;
            if (token.Length >= 2 && ((token[0] == '\'' && token[token.Length - 1] == '\'') || (token[0] == '"' && token[token.Length - 1] == '"'))) {
                quoted = true;
                token = token.Substring(1, token.Length - 2);
            }
            token = DecodeCssStringToken(token);
            if (token.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }
            var normalizedToken = token.Trim();
            var lowerToken = normalizedToken.ToLowerInvariant();
            return lowerToken switch {
                "disc" => "disc",
                "circle" => "circle",
                "square" => "square",
                "none" => "none",
                "decimal" => "decimal",
                "decimal-leading-zero" => "decimal-leading-zero",
                "lower-alpha" => "lower-alpha",
                "lower-latin" => "lower-alpha",
                "upper-alpha" => "upper-alpha",
                "upper-latin" => "upper-alpha",
                "lower-roman" => "lower-roman",
                "upper-roman" => "upper-roman",
                "lower-russian" => "lower-russian",
                "upper-russian" => "upper-russian",
                "hebrew" => "hebrew",
                "hebrew-1" => "hebrew-1",
                "hebrew-2" => "hebrew-2",
                "arabic-alpha" => "arabic-alpha",
                "arabic-abjad" => "arabic-abjad",
                "hiragana" => "hiragana",
                "hiragana-iroha" => "hiragana-iroha",
                "katakana" => "katakana",
                "katakana-iroha" => "katakana-iroha",
                "dash" => "dash",
                "hyphen" => "dash",
                "-" => "dash",
                "\u2013" => "en-dash",
                "en-dash" => "en-dash",
                "\u2014" => "em-dash",
                "em-dash" => "em-dash",
                "*" => "asterisk",
                "asterisk" => "asterisk",
                "+" => "plus",
                "plus" => "plus",
                _ => quoted && normalizedToken.Length > 0 ? "custom:" + normalizedToken : null,
            };
        }

        private static string DecodeCssStringToken(string token) {
            if (token.IndexOf('\\') < 0) {
                return token;
            }

            var decoded = new StringBuilder(token.Length);
            for (int i = 0; i < token.Length; i++) {
                if (token[i] != '\\' || i + 1 >= token.Length) {
                    decoded.Append(token[i]);
                    continue;
                }

                var start = i + 1;
                var end = start;
                while (end < token.Length && end - start < 6 && Uri.IsHexDigit(token[end])) {
                    end++;
                }

                if (end > start && int.TryParse(token.Substring(start, end - start), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var codePoint)) {
                    decoded.Append(char.ConvertFromUtf32(codePoint));
                    i = end - 1;
                    if (i + 1 < token.Length && char.IsWhiteSpace(token[i + 1])) {
                        i++;
                    }
                    continue;
                }

                decoded.Append(token[start]);
                i = start;
            }

            return decoded.ToString();
        }

        private static void ApplyListIndentMetadata(WordList list, int levelIndex, IElement element) {
            if (list.Numbering.Levels.Count <= levelIndex) {
                return;
            }

            var level = list.Numbering.Levels[levelIndex];
            if (TryGetTwipsAttribute(element, "data-left-indent-twips", out var leftIndentTwips)) {
                level.IndentationLeft = leftIndentTwips;
            }

            if (TryGetTwipsAttribute(element, "data-hanging-indent-twips", out var hangingIndentTwips)) {
                level.IndentationHanging = hangingIndentTwips;
            }
        }

        private static bool TryGetTwipsAttribute(IElement element, string name, out int value) {
            value = 0;
            var raw = element.GetAttribute(name);
            return !string.IsNullOrWhiteSpace(raw)
                && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
                && value >= 0;
        }

        private static void ApplyListStyle(WordList list, int levelIndex, bool ordered, string? listStyleType, string? typeAttr) {
            var levels = list.Numbering.Levels;
            if (levels.Count <= levelIndex) {
                return;
            }
            var level = levels[levelIndex];
            string? token = listStyleType;
            if (token != null) {
                token = token.Trim();
            }
            if (token == null || token.Length == 0) {
                if (typeAttr == null) {
                    return;
                }
                token = typeAttr.Trim();
                if (token.Length == 0) {
                    return;
                }
            }

            if (ordered) {
                var format = token switch {
                    "a" => NumberFormatValues.LowerLetter,
                    "A" => NumberFormatValues.UpperLetter,
                    "i" => NumberFormatValues.LowerRoman,
                    "I" => NumberFormatValues.UpperRoman,
                    "decimal-leading-zero" => NumberFormatValues.DecimalZero,
                    "lower-alpha" => NumberFormatValues.LowerLetter,
                    "upper-alpha" => NumberFormatValues.UpperLetter,
                    "lower-roman" => NumberFormatValues.LowerRoman,
                    "upper-roman" => NumberFormatValues.UpperRoman,
                    "lower-russian" => NumberFormatValues.RussianLower,
                    "upper-russian" => NumberFormatValues.RussianUpper,
                    "hebrew" => NumberFormatValues.Hebrew1,
                    "hebrew-1" => NumberFormatValues.Hebrew1,
                    "hebrew-2" => NumberFormatValues.Hebrew2,
                    "arabic-alpha" => NumberFormatValues.ArabicAlpha,
                    "arabic-abjad" => NumberFormatValues.ArabicAbjad,
                    "hiragana" => NumberFormatValues.Aiueo,
                    "hiragana-iroha" => NumberFormatValues.Iroha,
                    "katakana" => NumberFormatValues.AiueoFullWidth,
                    "katakana-iroha" => NumberFormatValues.IrohaFullWidth,
                    "none" => NumberFormatValues.None,
                    _ => NumberFormatValues.Decimal,
                };

                level._level!.NumberingFormat = new NumberingFormat { Val = format };
                if (format == NumberFormatValues.None) {
                    level.LevelText = string.Empty;
                }
                return;
            }

            var bulletToken = token.ToLowerInvariant();
            switch (bulletToken) {
                case "circle":
                    level.LevelText = "o";
                    break;
                case "square":
                    level.LevelText = "■";
                    break;
                case "none":
                    level._level.NumberingFormat = new NumberingFormat { Val = NumberFormatValues.None };
                    level.LevelText = string.Empty;
                    break;
                case "dash":
                    level.LevelText = "-";
                    break;
                case "en-dash":
                    level.LevelText = "\u2013";
                    break;
                case "em-dash":
                    level.LevelText = "\u2014";
                    break;
                case "asterisk":
                    level.LevelText = "*";
                    break;
                case "plus":
                    level.LevelText = "+";
                    break;
                default:
                    if (bulletToken.StartsWith("custom:", StringComparison.Ordinal)) {
                        level.LevelText = token.Substring("custom:".Length);
                    }
                    break;
                // disc/default -> no change
            }
        }

        private WordParagraph ProcessListItem(IHtmlListItemElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter,
            WordParagraph? insertionAnchor) {
            var list = listStack.Peek();
            int level = listStack.Count - 1;
            var paragraph = insertionAnchor == null
                ? list.AddItem("", level)
                : list.AddItemAfter("", level, insertionAnchor);
            WordParagraph? blockAnchor = paragraph;

            ApplyParagraphStyleFromCss(paragraph, element);
            ApplyClassStyle(element, paragraph, options);
            AddBookmarkIfPresent(element, paragraph);
            var bidi = GetBidiFromDir(element);
            if (bidi.HasValue) {
                paragraph.BiDi = bidi.Value;
            }
            foreach (var child in element.ChildNodes) {
                if (IsListItemTableChild(child)) {
                    var tableAnchor = blockAnchor ?? paragraph;
                    ProcessNode(child, doc, section, options, tableAnchor, listStack, formatting, cell, headerFooter);
                    blockAnchor = GetTrailingAnchorAfterTable(tableAnchor, section, cell, headerFooter);
                    continue;
                }

                var anchor = IsListItemBlockChild(child) ? blockAnchor : paragraph;
                ProcessNode(child, doc, section, options, anchor, listStack, formatting, cell, headerFooter);
                if (IsListItemBlockChild(child)) {
                    blockAnchor = GetLastParagraphInSameContainer(paragraph, section, cell, headerFooter);
                }
            }
            ApplyPageBreakAfterFromCss(paragraph, element);
            return blockAnchor ?? paragraph;
        }

        private static bool IsListItemTableChild(INode node) =>
            node is IElement element && string.Equals(element.TagName, "table", StringComparison.OrdinalIgnoreCase);

        private static bool IsListItemBlockChild(INode node) =>
            node is IElement element && _blockTags.Contains(element.TagName);

        private static WordParagraph GetLastParagraphInSameContainer(
            WordParagraph anchor,
            WordSection section,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter) {
            var parent = anchor._paragraph.Parent;
            var paragraphs = GetParagraphsInScope(section, cell, headerFooter);
            for (int i = paragraphs.Count - 1; i >= 0; i--) {
                var candidate = paragraphs[i];
                if (ReferenceEquals(candidate._paragraph.Parent, parent)) {
                    return candidate;
                }
            }

            return anchor;
        }

        private static WordParagraph? GetTrailingAnchorAfterTable(
            WordParagraph anchor,
            WordSection section,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter) {
            if (cell != null || headerFooter != null) {
                return GetLastParagraphInSameContainer(anchor, section, cell, headerFooter);
            }

            var parent = anchor._paragraph.Parent;
            if (parent == null) {
                return null;
            }

            var children = parent.ChildElements.ToList();
            var anchorIndex = children.IndexOf(anchor._paragraph);
            if (anchorIndex < 0) {
                return null;
            }

            var table = children
                .Skip(anchorIndex + 1)
                .OfType<Table>()
                .LastOrDefault();
            if (table == null) {
                return null;
            }

            var trailing = table.NextSibling<Paragraph>();
            if (trailing == null) {
                trailing = new Paragraph();
                table.InsertAfterSelf(trailing);
            }

            return new WordParagraph(anchor._document, trailing);
        }

        private static bool TryGetListItemValue(IHtmlListItemElement element, out int value) {
            value = 0;
            if (element == null) {
                return false;
            }

            var raw = element.GetAttribute("value");
            if (!string.IsNullOrWhiteSpace(raw) && int.TryParse(raw, out value)) {
                return true;
            }

            var attr = element.Attributes.FirstOrDefault(a => a.Name.Equals("value", StringComparison.OrdinalIgnoreCase));
            if (attr != null && !string.IsNullOrWhiteSpace(attr.Value) && int.TryParse(attr.Value, out value)) {
                return true;
            }

            if (element.Value.HasValue) {
                value = element.Value.Value;
                return true;
            }

            return false;
        }
    }
}
