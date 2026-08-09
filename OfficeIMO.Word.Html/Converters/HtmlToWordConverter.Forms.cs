using AngleSharp.Dom;
using OfficeIMO.Html;

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
                case "meter":
                case "progress":
                    ProcessValueElement(element, section, options, currentParagraph, formatting, cell, headerFooter);
                    break;
            }
        }

        private void ProcessInput(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            if (IsRadioInput(element)) {
                ProcessRadioGroup(element, section, options, currentParagraph, formatting, cell, headerFooter);
                return;
            }

            if (!IsCheckboxInput(element) && !IsTextInput(element) && !IsDateInput(element)) {
                return;
            }

            currentParagraph ??= AddParagraphInScope(section, cell, headerFooter);
            var (alias, tag) = GetInputMetadata(element);
            if (IsCheckboxInput(element)) {
                currentParagraph.AddCheckBox(IsCheckedInput(element), alias, tag);
            } else if (IsDateInput(element)) {
                var date = TryParseDateInput(HtmlFormControlSemantics.GetValues(element).FirstOrDefault());
                var datePicker = currentParagraph.AddDatePicker(date, alias, tag);
                datePicker.Date = date;
            } else if (TryGetDataListOptions(element, out var dataListOptions)) {
                var hasValueAttribute = element.HasAttribute("value");
                var value = HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty;
                string? selectedInternalValue = element.GetAttribute("data-word-value");
                int selectedIndex = selectedInternalValue == null
                    ? dataListOptions.FindIndex(option =>
                        string.Equals(option.DisplayText, value, StringComparison.Ordinal) ||
                        string.Equals(option.Value, value, StringComparison.Ordinal))
                    : dataListOptions.FindIndex(option =>
                        string.Equals(option.Value, selectedInternalValue, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(value) && selectedIndex < 0) {
                    dataListOptions.Insert(0, (value, value));
                    selectedIndex = 0;
                }
                if (!hasValueAttribute) {
                    selectedIndex = dataListOptions.FindIndex(option => option.DisplayText.Length == 0);
                    if (selectedIndex < 0) selectedIndex = 0;
                }
                if (selectedIndex < 0) selectedIndex = 0;
                var comboBox = currentParagraph.AddComboBox(
                    dataListOptions.Select(option => option.DisplayText).ToList(),
                    alias,
                    tag,
                    dataListOptions[selectedIndex].DisplayText);
                comboBox.SetImportedItems(dataListOptions, selectedIndex);
            } else {
                string value = HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty;
                currentParagraph.AddStructuredDocumentTag(value, alias, tag);
            }

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private void ProcessRadioGroup(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            if (_processedRadioInputs.Contains(element)) {
                return;
            }

            var group = GetRadioGroup(element);
            if (group.Count == 0) {
                return;
            }

            foreach (var radio in group) {
                _processedRadioInputs.Add(radio);
            }

            var optionTexts = group.Select(GetRadioOptionText).ToList();
            if (optionTexts.Count == 0) {
                return;
            }

            var selected = group
                .Where(IsCheckedInput)
                .Select(GetRadioOptionText)
                .FirstOrDefault();

            if (selected == null && !optionTexts.Contains(string.Empty, StringComparer.Ordinal)) {
                optionTexts.Insert(0, string.Empty);
                selected = string.Empty;
            }

            currentParagraph ??= AddParagraphInScope(section, cell, headerFooter);
            var (alias, tag) = GetRadioGroupMetadata(group);
            var dropDown = currentParagraph.AddDropDownList(optionTexts, alias, tag);
            dropDown.SelectedValue = selected ?? string.Empty;

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private void ProcessSelect(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            var optionsList = element.QuerySelectorAll("option")
                .Select(option => new {
                    Value = GetOptionValue(option),
                    DisplayText = NormalizeFormText(option.TextContent),
                    Selected = option.HasAttribute("selected")
                })
                .ToList();

            if (optionsList.Count == 0) {
                return;
            }

            currentParagraph ??= AddParagraphInScope(section, cell, headerFooter);
            var (alias, tag) = GetInputMetadata(element);
            if (element.HasAttribute("multiple")) {
                var selectedValues = optionsList
                    .Where(option => option.Selected)
                    .Select(option => option.Value)
                    .ToList();
                currentParagraph.AddStructuredDocumentTag(string.Join("\n", selectedValues), alias, tag);
                if (ShouldAddSpaceAfterInput(element)) {
                    AddTextRun(currentParagraph, " ", formatting, options);
                }

                return;
            }

            int selectedIndex = optionsList.FindLastIndex(option => option.Selected);
            if (selectedIndex < 0) selectedIndex = 0;
            var importedItems = optionsList
                .Select(option => (option.Value, option.DisplayText))
                .ToList();
            var dropDown = currentParagraph.AddDropDownList(
                importedItems.Select(option => option.DisplayText), alias, tag);
            dropDown.SetImportedItems(importedItems, selectedIndex);

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private void ProcessTextArea(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            currentParagraph ??= AddParagraphInScope(section, cell, headerFooter);
            var (alias, tag) = GetInputMetadata(element);
            currentParagraph.AddStructuredDocumentTag(NormalizeFormText(element.TextContent), alias, tag);

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private void ProcessValueElement(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            currentParagraph ??= AddParagraphInScope(section, cell, headerFooter);
            var (alias, tag) = GetInputMetadata(element);
            currentParagraph.AddStructuredDocumentTag(GetValueElementText(element), alias, tag);

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private static bool IsCheckboxInput(IElement element) {
            return string.Equals(GetEffectiveInputType(element), "checkbox", StringComparison.Ordinal);
        }

        private static bool IsRadioInput(IElement element) {
            return string.Equals(GetEffectiveInputType(element), "radio", StringComparison.Ordinal);
        }

        private static bool IsCheckedInput(IElement element) =>
            HtmlFormControlSemantics.IsEffectivelyChecked(element);

        private static bool IsDateInput(IElement element) {
            return string.Equals(GetEffectiveInputType(element), "date", StringComparison.Ordinal);
        }

        private static bool IsTextInput(IElement element) {
            return GetEffectiveInputType(element) switch {
                "text" or "search" or "email" or "url" or "tel" or "password" or
                "number" or "time" or "datetime-local" or "month" or "week" or "color" or "range" => true,
                _ => false,
            };
        }

        private static string GetEffectiveInputType(IElement element) =>
            HtmlFormControlSemantics.GetEffectiveType("input", element.GetAttribute("type"));

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

        private static string GetValueElementText(IElement element) {
            var value = element.GetAttribute("value");
            if (string.IsNullOrWhiteSpace(value)) {
                return NormalizeFormText(element.TextContent).Trim();
            }

            var max = element.GetAttribute("max");
            return string.IsNullOrWhiteSpace(max) ? value! : $"{value} / {max}";
        }

        private static string GetOptionValue(IElement option) =>
            NormalizeFormText(option.GetAttribute("value") ?? option.TextContent);

        private static List<IElement> GetRadioGroup(IElement element) {
            var name = element.GetAttribute("name");
            if (string.IsNullOrEmpty(name)) {
                return new List<IElement> { element };
            }

            var root = GetRootElement(element);
            var formOwner = HtmlFormControlSemantics.ResolveFormOwner(element);
            return root.QuerySelectorAll("input")
                .Where(IsRadioInput)
                .Where(input => string.Equals(input.GetAttribute("name"), name, StringComparison.Ordinal))
                .Where(input => ReferenceEquals(HtmlFormControlSemantics.ResolveFormOwner(input), formOwner))
                .ToList();
        }

        private static string GetRadioOptionText(IElement element) {
            var value = element.GetAttribute("value");
            if (!string.IsNullOrEmpty(value)) {
                return NormalizeFormText(value);
            }

            var label = GetRadioLabelText(element);
            if (!string.IsNullOrWhiteSpace(label)) {
                return label!;
            }

            return NormalizeFormText(element.GetAttribute("aria-label") ?? element.GetAttribute("title") ?? element.GetAttribute("id") ?? element.GetAttribute("name") ?? "on");
        }

        private static (string? Alias, string? Tag) GetRadioGroupMetadata(IReadOnlyList<IElement> group) {
            var checkedInput = group.FirstOrDefault(IsCheckedInput);
            var first = group[0];
            var metadataSource = checkedInput ?? first;
            var name = first.GetAttribute("name");
            var alias = metadataSource.GetAttribute("aria-label") ?? metadataSource.GetAttribute("title") ?? name ?? metadataSource.GetAttribute("id");
            var tag = metadataSource.GetAttribute("data-tag") ?? name ?? metadataSource.GetAttribute("id");
            return (alias, tag);
        }

        private static bool IsRadioChoiceLabel(IElement element) {
            if (!string.Equals(element.TagName, "label", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (element.QuerySelectorAll("input").Any(IsRadioInput)) {
                return true;
            }

            var targetId = element.GetAttribute("for");
            if (string.IsNullOrWhiteSpace(targetId)) {
                return false;
            }

            var target = FindElementById(GetRootElement(element), targetId!);
            return target != null && IsRadioInput(target);
        }

        private static string? GetRadioLabelText(IElement element) {
            var current = element.ParentElement;
            while (current != null) {
                if (string.Equals(current.TagName, "label", StringComparison.OrdinalIgnoreCase)) {
                    return NormalizeFormText(current.TextContent).Trim();
                }

                current = current.ParentElement;
            }

            var id = element.GetAttribute("id");
            if (string.IsNullOrWhiteSpace(id)) {
                return null;
            }

            var root = GetRootElement(element);
            var labels = root.QuerySelectorAll("label")
                .Where(label => string.Equals(label.GetAttribute("for"), id, StringComparison.Ordinal))
                .Select(label => NormalizeFormText(label.TextContent).Trim())
                .Where(text => text.Length > 0)
                .ToList();

            return labels.Count == 0 ? null : labels[0];
        }

        private static IElement GetRootElement(IElement element) {
            var root = element;
            while (root.ParentElement != null) {
                root = root.ParentElement;
            }

            return root;
        }

        private static IElement? FindElementById(IElement root, string id) {
            var stack = new Stack<IElement>();
            stack.Push(root);
            while (stack.Count > 0) {
                var current = stack.Pop();
                if (string.Equals(current.GetAttribute("id"), id, StringComparison.Ordinal)) {
                    return current;
                }

                foreach (var child in current.Children) {
                    stack.Push(child);
                }
            }

            return null;
        }

        private static bool TryGetDataListOptions(
            IElement element,
            out List<(string Value, string DisplayText)> options) {
            options = new List<(string Value, string DisplayText)>();
            var listId = element.GetAttribute("list");
            if (string.IsNullOrWhiteSpace(listId)) {
                return false;
            }

            var root = element;
            while (root.ParentElement != null) {
                root = root.ParentElement;
            }

            var dataList = FindDataListElement(root, listId!);
            if (dataList == null) {
                return false;
            }

            options = dataList.QuerySelectorAll("option")
                .Select(option => {
                    string displayText = GetOptionValue(option);
                    string value = option.GetAttribute("data-word-value") ?? displayText;
                    return (value, displayText);
                })
                .ToList();

            return options.Count > 0;
        }

        private static IElement? FindDataListElement(IElement root, string listId) {
            var stack = new Stack<IElement>();
            stack.Push(root);
            while (stack.Count > 0) {
                var current = stack.Pop();
                if (string.Equals(current.TagName, "datalist", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(current.GetAttribute("id"), listId, StringComparison.Ordinal)) {
                    return current;
                }

                foreach (var child in current.Children) {
                    stack.Push(child);
                }
            }

            return null;
        }

        private static bool ShouldAddSpaceAfterInput(IElement element) {
            var sibling = element.NextSibling;
            while (sibling is IElement siblingElement &&
                string.Equals(siblingElement.TagName, "datalist", StringComparison.OrdinalIgnoreCase)) {
                sibling = sibling.NextSibling;
            }

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
