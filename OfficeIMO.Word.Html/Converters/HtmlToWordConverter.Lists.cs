using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Wordprocessing;

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
            ApplyListStyle(list, ordered, listStyleType, typeAttr);

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
                    list.SetStartNumberingValue(startValue.Value);
                }
            }

            listStack.Push(list);

            int itemIndex = 0;
            foreach (var li in element.Children.OfType<IHtmlListItemElement>()) {
                if (ordered) {
                    if (TryGetListItemValue(li, out int liValue)) {
                        if (itemIndex == 0) {
                            list.SetStartNumberingValue(liValue);
                        } else {
                            list = CreateOrderedList(allowContinue: false);
                            ApplyListStyle(list, ordered, listStyleType, typeAttr);
                            list.SetStartNumberingValue(liValue);
                            listStack.Pop();
                            listStack.Push(list);
                        }
                    }
                }

                ProcessListItem(li, doc, section, options, listStack, formatting, cell, headerFooter);
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
            if (token.Length >= 2 && ((token[0] == '\'' && token[token.Length - 1] == '\'') || (token[0] == '"' && token[token.Length - 1] == '"'))) {
                token = token.Substring(1, token.Length - 2);
            }
            if (token.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }
            token = token.Trim().ToLowerInvariant();
            return token switch {
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
                _ => null,
            };
        }

        private static void ApplyListStyle(WordList list, bool ordered, string? listStyleType, string? typeAttr) {
            var levels = list.Numbering.Levels;
            if (levels.Count == 0) {
                return;
            }
            var level = levels[0];
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
                // disc/default -> no change
            }
        }

        private void ProcessListItem(IHtmlListItemElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            var list = listStack.Peek();
            int level = listStack.Count - 1;
            var paragraph = list.AddItem("", level);

            ApplyParagraphStyleFromCss(paragraph, element);
            ApplyClassStyle(element, paragraph, options);
            AddBookmarkIfPresent(element, paragraph);
            var bidi = GetBidiFromDir(element);
            if (bidi.HasValue) {
                paragraph.BiDi = bidi.Value;
            }
            foreach (var child in element.ChildNodes) {
                ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell, headerFooter);
            }
            ApplyPageBreakAfterFromCss(paragraph, element);
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
