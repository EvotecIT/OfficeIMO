using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ProcessFormControl(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            switch (element.TagName.ToLowerInvariant()) {
                case "input":
                    ProcessInput(element, section, options, currentParagraph, formatting, cell, headerFooter);
                    break;
                case "select":
                    ProcessSelect(element, section, options, currentParagraph, formatting, cell, headerFooter);
                    break;
                case "textarea":
                    ProcessTextArea(element, section, options, currentParagraph, formatting, cell, headerFooter);
                    break;
            }
        }

        private void ProcessInput(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            if (!IsCheckboxInput(element) && !IsTextInput(element) && !IsDateInput(element)) {
                return;
            }

            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            var (alias, tag) = GetInputMetadata(element);
            if (IsCheckboxInput(element)) {
                currentParagraph.AddCheckBox(IsCheckedInput(element), alias, tag);
            } else if (IsDateInput(element)) {
                currentParagraph.AddDatePicker(TryParseDateInput(element.GetAttribute("value")), alias, tag);
            } else {
                currentParagraph.AddStructuredDocumentTag(element.GetAttribute("value") ?? string.Empty, alias, tag);
            }

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private void ProcessSelect(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            var optionsList = element.Children
                .Where(child => string.Equals(child.TagName, "option", StringComparison.OrdinalIgnoreCase))
                .Select(option => new {
                    Text = NormalizeFormText(option.TextContent),
                    Selected = option.HasAttribute("selected")
                })
                .Where(option => option.Text.Length > 0)
                .ToList();

            if (optionsList.Count == 0) {
                return;
            }

            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            var (alias, tag) = GetInputMetadata(element);
            var dropDown = currentParagraph.AddDropDownList(optionsList.Select(option => option.Text), alias, tag);
            var selected = optionsList.FirstOrDefault(option => option.Selected)?.Text ?? optionsList[0].Text;
            dropDown.SelectedValue = selected;

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private void ProcessTextArea(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            var (alias, tag) = GetInputMetadata(element);
            currentParagraph.AddStructuredDocumentTag(NormalizeFormText(element.TextContent), alias, tag);

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private static bool IsCheckboxInput(IElement element) {
            var type = element.GetAttribute("type");
            return string.Equals(type, "checkbox", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCheckedInput(IElement element) =>
            element.HasAttribute("checked") ||
            string.Equals(element.GetAttribute("aria-checked"), "true", StringComparison.OrdinalIgnoreCase);

        private static bool IsDateInput(IElement element) {
            var type = element.GetAttribute("type");
            return string.Equals(type, "date", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTextInput(IElement element) {
            var type = element.GetAttribute("type");
            if (string.IsNullOrWhiteSpace(type)) {
                return true;
            }

            var normalizedType = type!.ToLowerInvariant();
            return normalizedType switch {
                "text" or "search" or "email" or "url" or "tel" or "password" => true,
                _ => false,
            };
        }

        private static (string? Alias, string? Tag) GetInputMetadata(IElement element) {
            var id = element.GetAttribute("id");
            var name = element.GetAttribute("name");
            var alias = element.GetAttribute("aria-label") ?? element.GetAttribute("title") ?? name ?? id;
            var dataTag = element.GetAttribute("data-tag");
            var tag = dataTag ?? id ?? name;
            return (alias, tag);
        }

        private static DateTime? TryParseDateInput(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (DateTime.TryParseExact(value, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var date)) {
                return date;
            }

            return null;
        }

        private static string NormalizeFormText(string? text) =>
            text?.Replace("\r\n", "\n").Replace('\r', '\n') ?? string.Empty;

        private static bool ShouldAddSpaceAfterInput(IElement element) {
            var sibling = element.NextSibling;
            if (sibling == null) {
                return false;
            }

            if (sibling is IText text) {
                return text.Text.Length > 0 && !char.IsWhiteSpace(text.Text[0]);
            }

            return true;
        }
    }
}
