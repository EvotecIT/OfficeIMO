using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        IElement CreateCheckBoxInput(IDocument htmlDoc, WordCheckBox checkBox) {
            var input = htmlDoc.CreateElement("input");
            input.SetAttribute("type", "checkbox");
            input.SetAttribute("disabled", string.Empty);

            if (checkBox.IsChecked) {
                input.SetAttribute("checked", string.Empty);
            }

            ApplyContentControlMetadata(input, checkBox.Alias, checkBox.Tag);

            return input;
        }

        IElement CreateDropDownListSelect(IDocument htmlDoc, WordDropDownList dropDownList) {
            var select = htmlDoc.CreateElement("select");
            select.SetAttribute("disabled", string.Empty);
            ApplyContentControlMetadata(select, dropDownList.Alias, dropDownList.Tag);

            foreach (var item in dropDownList.Items) {
                var option = htmlDoc.CreateElement("option");
                option.SetAttribute("value", item);
                option.TextContent = item;

                if (string.Equals(item, dropDownList.SelectedValue, StringComparison.OrdinalIgnoreCase)) {
                    option.SetAttribute("selected", string.Empty);
                }

                select.AppendChild(option);
            }

            return select;
        }

        IEnumerable<INode> CreateComboBoxNodes(IDocument htmlDoc, WordComboBox comboBox, int formListIndex) {
            string listId = "word-combo-" + formListIndex.ToString(CultureInfo.InvariantCulture);

            var input = htmlDoc.CreateElement("input");
            input.SetAttribute("type", "text");
            input.SetAttribute("disabled", string.Empty);
            input.SetAttribute("list", listId);
            if (!string.IsNullOrEmpty(comboBox.SelectedValue)) {
                input.SetAttribute("value", comboBox.SelectedValue!);
            }
            ApplyContentControlMetadata(input, comboBox.Alias, comboBox.Tag);
            yield return input;

            var dataList = htmlDoc.CreateElement("datalist");
            dataList.SetAttribute("id", listId);
            foreach (var item in comboBox.Items) {
                var option = htmlDoc.CreateElement("option");
                option.SetAttribute("value", item);
                dataList.AppendChild(option);
            }

            yield return dataList;
        }

        IElement CreateDatePickerInput(IDocument htmlDoc, WordDatePicker datePicker) {
            var input = htmlDoc.CreateElement("input");
            input.SetAttribute("type", "date");
            input.SetAttribute("disabled", string.Empty);
            if (datePicker.Date.HasValue) {
                input.SetAttribute("value", datePicker.Date.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            }
            ApplyContentControlMetadata(input, datePicker.Alias, datePicker.Tag);
            return input;
        }

        IElement CreateStructuredDocumentTagInput(IDocument htmlDoc, WordStructuredDocumentTag structuredDocumentTag) {
            var input = htmlDoc.CreateElement("input");
            input.SetAttribute("type", "text");
            input.SetAttribute("disabled", string.Empty);
            if (!string.IsNullOrEmpty(structuredDocumentTag.Text)) {
                input.SetAttribute("value", structuredDocumentTag.Text!);
            }
            ApplyContentControlMetadata(input, structuredDocumentTag.Alias, structuredDocumentTag.Tag);
            return input;
        }

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
