using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        IElement CreateCheckBoxInput(IDocument htmlDoc, WordCheckBox checkBox) {
            var input = CreateOutputElement(htmlDoc, "input");
            input.SetAttribute("type", "checkbox");
            input.SetAttribute("disabled", string.Empty);

            if (checkBox.IsChecked) {
                input.SetAttribute("checked", string.Empty);
            }

            ApplyContentControlMetadata(input, checkBox.Alias, checkBox.Tag);

            return input;
        }

        IElement CreateDropDownListSelect(IDocument htmlDoc, WordDropDownList dropDownList) {
            var select = CreateOutputElement(htmlDoc, "select");
            select.SetAttribute("disabled", string.Empty);
            ApplyContentControlMetadata(select, dropDownList.Alias, dropDownList.Tag);

            foreach (var item in dropDownList.ExportItems) {
                var option = CreateOutputElement(htmlDoc, "option");
                SetOutputAttribute(htmlDoc, option, "value", item.Value, "DropDownOption:value");
                SetOutputText(htmlDoc, option, item.DisplayText, "DropDownOption:display-text");

                if (string.Equals(item.Value, dropDownList.SelectedValue, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.DisplayText, dropDownList.SelectedValue, StringComparison.OrdinalIgnoreCase)) {
                    option.SetAttribute("selected", string.Empty);
                }

                select.AppendChild(option);
            }

            return select;
        }

        IEnumerable<INode> CreateComboBoxNodes(IDocument htmlDoc, WordComboBox comboBox, int formListIndex) {
            string listId = "word-combo-" + formListIndex.ToString(CultureInfo.InvariantCulture);

            var input = CreateOutputElement(htmlDoc, "input");
            input.SetAttribute("type", "text");
            input.SetAttribute("disabled", string.Empty);
            input.SetAttribute("list", listId);
            if (!string.IsNullOrEmpty(comboBox.SelectedValue)) {
                input.SetAttribute("value", comboBox.SelectedValue!);
            }
            ApplyContentControlMetadata(input, comboBox.Alias, comboBox.Tag);
            yield return input;

            var dataList = CreateOutputElement(htmlDoc, "datalist");
            dataList.SetAttribute("id", listId);
            foreach (var item in comboBox.ExportItems) {
                var option = CreateOutputElement(htmlDoc, "option");
                SetOutputAttribute(htmlDoc, option, "value", item.Value, "ComboBoxOption:value");
                dataList.AppendChild(option);
            }

            yield return dataList;
        }

        IElement CreateDatePickerInput(IDocument htmlDoc, WordDatePicker datePicker) {
            var input = CreateOutputElement(htmlDoc, "input");
            input.SetAttribute("type", "date");
            input.SetAttribute("disabled", string.Empty);
            if (datePicker.Date.HasValue) {
                input.SetAttribute("value", datePicker.Date.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            }
            ApplyContentControlMetadata(input, datePicker.Alias, datePicker.Tag);
            return input;
        }

        IElement CreateStructuredDocumentTagInput(IDocument htmlDoc, WordStructuredDocumentTag structuredDocumentTag) {
            if (HasLineBreaks(structuredDocumentTag.Text)) {
                var textArea = CreateOutputElement(htmlDoc, "textarea");
                textArea.SetAttribute("disabled", string.Empty);
                textArea.TextContent = structuredDocumentTag.Text ?? string.Empty;
                ApplyContentControlMetadata(textArea, structuredDocumentTag.Alias, structuredDocumentTag.Tag);
                return textArea;
            }

            var input = CreateOutputElement(htmlDoc, "input");
            input.SetAttribute("type", "text");
            input.SetAttribute("disabled", string.Empty);
            if (!string.IsNullOrEmpty(structuredDocumentTag.Text)) {
                input.SetAttribute("value", structuredDocumentTag.Text!);
            }
            ApplyContentControlMetadata(input, structuredDocumentTag.Alias, structuredDocumentTag.Tag);
            return input;
        }

        static bool HasLineBreaks(string? text) =>
            !string.IsNullOrEmpty(text) && (text!.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0);

        static void ApplyContentControlMetadata(IElement element, string? alias, string? tag) {
            if (!string.IsNullOrEmpty(alias)) {
                element.SetAttribute("aria-label", alias!);
            }

            if (!string.IsNullOrEmpty(tag)) {
                element.SetAttribute("data-tag", tag!);
            }
        }
    }
}
