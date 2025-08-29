using OfficeIMO.Word;
using System.Collections.Generic;
                    list.NumberId = _orderedListNumberId.Value;
                        _orderedListNumberId = list.NumberId;

using System.Collections.Generic;
using System.Reflection;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private int? _orderedListNumberId;

        private void ProcessList(IElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, WordTableCell? cell, TextFormatting formatting, WordHeaderFooter? headerFooter) {
            WordList list;
            bool ordered = element.TagName.Equals("ol", System.StringComparison.OrdinalIgnoreCase);
            if (ordered) {
                if (options.ContinueNumbering && _orderedListNumberId.HasValue && listStack.Count == 0 && cell == null && headerFooter == null) {
                    list = new WordList(doc);
                    var field = typeof(WordList).GetField("_numberId", BindingFlags.NonPublic | BindingFlags.Instance);
                    field?.SetValue(list, _orderedListNumberId.Value);
                } else {
                    // Use standard numbered list style for ordered lists in all contexts
                    list = cell != null ? cell.AddList(WordListStyle.Numbered) : headerFooter != null ? headerFooter.AddList(WordListStyle.Numbered) : doc.AddListNumbered();
                    if (options.ContinueNumbering && listStack.Count == 0 && cell == null && headerFooter == null) {
                        var field = typeof(WordList).GetField("_numberId", BindingFlags.NonPublic | BindingFlags.Instance);
                        _orderedListNumberId = (int?)field?.GetValue(list);
                    }
                }
                var level = list.Numbering.Levels[0];
                var start = element.GetAttribute("start");
                if (!string.IsNullOrEmpty(start) && int.TryParse(start, out int startVal)) {
                    level.SetStartNumberingValue(startVal);
                }
                var type = element.GetAttribute("type");
                if (!string.IsNullOrEmpty(type)) {
                    var format = type switch {
                        "a" => NumberFormatValues.LowerLetter,
                        "A" => NumberFormatValues.UpperLetter,
                        "i" => NumberFormatValues.LowerRoman,
                        "I" => NumberFormatValues.UpperRoman,
                        _ => NumberFormatValues.Decimal,
                    };
                    level._level.NumberingFormat = new NumberingFormat { Val = format };
                }
            } else {
                list = cell != null ? cell.AddList(WordListStyle.Bulleted) : headerFooter != null ? headerFooter.AddList(WordListStyle.Bulleted) : doc.AddListBulleted();
                var type = element.GetAttribute("type")?.ToLowerInvariant();
                if (!string.IsNullOrEmpty(type)) {
                    var level = list.Numbering.Levels[0];
                    switch (type) {
                        case "circle":
                            level._level.LevelText.Val = "o";
                            break;
                        case "square":
                            level._level.LevelText.Val = "■";
                            break;
                        // disc is the default, nothing to change
                    }
                }
            }
            listStack.Push(list);
            foreach (var li in element.Children.OfType<IHtmlListItemElement>()) {
                ProcessListItem(li, doc, section, options, listStack, formatting, cell, headerFooter);
            }
            listStack.Pop();
        }

        private void ProcessListItem(IHtmlListItemElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            var list = listStack.Peek();
            int level = listStack.Count - 1;
            var paragraph = list.AddItem("", level);
            ApplyClassStyle(element, paragraph, options);
            AddBookmarkIfPresent(element, paragraph);
            foreach (var child in element.ChildNodes) {
                ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell, headerFooter);
            }
        }
    }
}
