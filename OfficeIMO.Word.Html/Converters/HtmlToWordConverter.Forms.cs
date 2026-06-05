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

            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            var (alias, tag) = GetInputMetadata(element);
            if (IsCheckboxInput(element)) {
                currentParagraph.AddCheckBox(IsCheckedInput(element), alias, tag);
            } else if (IsDateInput(element)) {
                var date = TryParseDateInput(element.GetAttribute("value"));
                var datePicker = currentParagraph.AddDatePicker(date, alias, tag);
                datePicker.Date = date;
            } else if (TryGetDataListOptions(element, out var dataListOptions)) {
                var hasValueAttribute = element.HasAttribute("value");
                var value = element.GetAttribute("value") ?? string.Empty;
                if (!string.IsNullOrEmpty(value) && !dataListOptions.Contains(value, StringComparer.Ordinal)) {
                    dataListOptions.Insert(0, value);
                }
                var defaultValue = dataListOptions.Contains(value, StringComparer.Ordinal) ? value : null;
                var comboBox = currentParagraph.AddComboBox(dataListOptions, alias, tag, defaultValue);
                if (!hasValueAttribute && dataListOptions.Contains(string.Empty, StringComparer.Ordinal)) {
                    comboBox.SelectedValue = string.Empty;
                }
            } else {
                currentParagraph.AddStructuredDocumentTag(element.GetAttribute("value") ?? string.Empty, alias, tag);
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

            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
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
                    Text = GetOptionText(option),
                    Selected = option.HasAttribute("selected")
                })
                .ToList();

            if (optionsList.Count == 0) {
                return;
            }

            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            var (alias, tag) = GetInputMetadata(element);
            if (element.HasAttribute("multiple")) {
                var selectedValues = optionsList
                    .Where(option => option.Selected)
                    .Select(option => option.Text)
                    .ToList();
                currentParagraph.AddStructuredDocumentTag(string.Join("\n", selectedValues), alias, tag);
                if (ShouldAddSpaceAfterInput(element)) {
                    AddTextRun(currentParagraph, " ", formatting, options);
                }

                return;
            }

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

        private void ProcessValueElement(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            var (alias, tag) = GetInputMetadata(element);
            currentParagraph.AddStructuredDocumentTag(GetValueElementText(element), alias, tag);

            if (ShouldAddSpaceAfterInput(element)) {
                AddTextRun(currentParagraph, " ", formatting, options);
            }
        }

        private static bool IsCheckboxInput(IElement element) {
            var type = element.GetAttribute("type");
            return string.Equals(type, "checkbox", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRadioInput(IElement element) {
            var type = element.GetAttribute("type");
            return string.Equals(type, "radio", StringComparison.OrdinalIgnoreCase);
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
                "text" or "search" or "email" or "url" or "tel" or "password" or
                "number" or "time" or "datetime-local" or "month" or "week" or "color" or "range" => true,
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

        private static string GetValueElementText(IElement element) {
            var value = element.GetAttribute("value");
            if (string.IsNullOrWhiteSpace(value)) {
                return NormalizeFormText(element.TextContent).Trim();
            }

            var max = element.GetAttribute("max");
            return string.IsNullOrWhiteSpace(max) ? value! : $"{value} / {max}";
        }

        private static string GetOptionText(IElement option) =>
            NormalizeFormText(option.GetAttribute("value") ?? option.TextContent);

        private static List<IElement> GetRadioGroup(IElement element) {
            var name = element.GetAttribute("name");
            if (string.IsNullOrWhiteSpace(name)) {
                return new List<IElement> { element };
            }

            var root = GetRootElement(element);
            var explicitFormOwner = element.GetAttribute("form");
            var ancestorFormOwner = FindAncestorForm(element);
            return root.QuerySelectorAll("input")
                .Where(IsRadioInput)
                .Where(input => string.Equals(input.GetAttribute("name"), name, StringComparison.Ordinal))
                .Where(input => SameRadioFormOwner(input, explicitFormOwner, ancestorFormOwner))
                .ToList();
        }

        private static bool SameRadioFormOwner(IElement element, string? explicitFormOwner, IElement? ancestorFormOwner) {
            var elementExplicitFormOwner = element.GetAttribute("form");
            if (!string.IsNullOrWhiteSpace(explicitFormOwner) || !string.IsNullOrWhiteSpace(elementExplicitFormOwner)) {
                return string.Equals(elementExplicitFormOwner, explicitFormOwner, StringComparison.Ordinal);
            }

            return ReferenceEquals(FindAncestorForm(element), ancestorFormOwner);
        }

        private static IElement? FindAncestorForm(IElement element) {
            var current = element.ParentElement;
            while (current != null) {
                if (string.Equals(current.TagName, "form", StringComparison.OrdinalIgnoreCase)) {
                    return current;
                }

                current = current.ParentElement;
            }

            return null;
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

        private static bool TryGetDataListOptions(IElement element, out List<string> options) {
            options = new List<string>();
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
                .Select(GetOptionText)
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
