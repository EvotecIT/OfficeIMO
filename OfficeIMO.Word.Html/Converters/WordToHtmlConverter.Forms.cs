using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        IElement CreateCheckBoxInput(IDocument htmlDoc, WordCheckBox checkBox) {
            ReleaseReplacedContentControlContent(htmlDoc, checkBox._sdtRun);
            var input = CreateOutputElement(htmlDoc, "input");
            SetOutputAttribute(input, "type", "checkbox", "CheckBox:type");
            SetOutputAttribute(input, "disabled", string.Empty, "CheckBox:disabled");

            if (checkBox.IsChecked) {
                SetOutputAttribute(input, "checked", string.Empty, "CheckBox:checked");
            }

            ApplyContentControlMetadata(input, checkBox.Alias, checkBox.Tag);

            return input;
        }

        IElement CreateDropDownListSelect(IDocument htmlDoc, WordDropDownList dropDownList) {
            ReleaseReplacedContentControlContent(htmlDoc, dropDownList._sdtRun);
            var select = CreateOutputElement(htmlDoc, "select");
            SetOutputAttribute(select, "disabled", string.Empty, "DropDown:disabled");
            ApplyContentControlMetadata(select, dropDownList.Alias, dropDownList.Tag);

            foreach (var item in dropDownList.ExportItems) {
                var option = CreateOutputElement(htmlDoc, "option");
                SetOutputAttribute(htmlDoc, option, "value", item.Value, "DropDownOption:value");
                SetOutputText(htmlDoc, option, item.DisplayText, "DropDownOption:display-text");

                if (string.Equals(item.Value, dropDownList.SelectedValue, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.DisplayText, dropDownList.SelectedValue, StringComparison.OrdinalIgnoreCase)) {
                    SetOutputAttribute(option, "selected", string.Empty, "DropDownOption:selected");
                }

                select.AppendChild(option);
            }

            return select;
        }

        IEnumerable<INode> CreateComboBoxNodes(IDocument htmlDoc, WordComboBox comboBox, int formListIndex) {
            ReleaseReplacedContentControlContent(htmlDoc, comboBox._sdtRun);
            string listId = "word-combo-" + formListIndex.ToString(CultureInfo.InvariantCulture);
            IReadOnlyList<(string Value, string DisplayText)> items = comboBox.ExportItems;
            string? selectedValue = comboBox.SelectedValue;
            string? selectedDisplayText = items
                .FirstOrDefault(item =>
                    string.Equals(item.Value, selectedValue, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.DisplayText, selectedValue, StringComparison.OrdinalIgnoreCase))
                .DisplayText;

            var input = CreateOutputElement(htmlDoc, "input");
            SetOutputAttribute(input, "type", "text", "ComboBox:type");
            SetOutputAttribute(input, "disabled", string.Empty, "ComboBox:disabled");
            SetOutputAttribute(input, "list", listId, "ComboBox:list");
            if (!string.IsNullOrEmpty(selectedValue)) {
                SetOutputAttribute(
                    htmlDoc,
                    input,
                    "value",
                    string.IsNullOrEmpty(selectedDisplayText) ? selectedValue! : selectedDisplayText!,
                    "ComboBox:selected-display");
            }
            ApplyContentControlMetadata(input, comboBox.Alias, comboBox.Tag);
            yield return input;

            var dataList = CreateOutputElement(htmlDoc, "datalist");
            SetOutputAttribute(dataList, "id", listId, "ComboBoxList:id");
            foreach (var item in items) {
                var option = CreateOutputElement(htmlDoc, "option");
                SetOutputAttribute(htmlDoc, option, "value", item.DisplayText, "ComboBoxOption:value");
                SetOutputAttribute(htmlDoc, option, "label", item.DisplayText, "ComboBoxOption:label");
                if (!string.Equals(item.Value, item.DisplayText, StringComparison.Ordinal)) {
                    SetOutputAttribute(htmlDoc, option, "data-word-value", item.Value, "ComboBoxOption:internal-value");
                }
                dataList.AppendChild(option);
            }

            yield return dataList;
        }

        IElement CreateDatePickerInput(IDocument htmlDoc, WordDatePicker datePicker) {
            ReleaseReplacedContentControlContent(htmlDoc, datePicker._sdtRun);
            var input = CreateOutputElement(htmlDoc, "input");
            SetOutputAttribute(input, "type", "date", "DatePicker:type");
            SetOutputAttribute(input, "disabled", string.Empty, "DatePicker:disabled");
            if (datePicker.Date.HasValue) {
                SetOutputAttribute(input, "value", datePicker.Date.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), "DatePicker:value");
            }
            ApplyContentControlMetadata(input, datePicker.Alias, datePicker.Tag);
            return input;
        }

        IElement CreateStructuredDocumentTagInput(IDocument htmlDoc, WordStructuredDocumentTag structuredDocumentTag) {
            ReleaseReplacedContentControlContent(htmlDoc, structuredDocumentTag.SdtElement);
            if (HasLineBreaks(structuredDocumentTag.Text)) {
                var textArea = CreateOutputElement(htmlDoc, "textarea");
                SetOutputAttribute(textArea, "disabled", string.Empty, "ContentControl:disabled");
                SetOutputText(htmlDoc, textArea, structuredDocumentTag.Text ?? string.Empty, "ContentControl:text");
                ApplyContentControlMetadata(textArea, structuredDocumentTag.Alias, structuredDocumentTag.Tag);
                return textArea;
            }

            var input = CreateOutputElement(htmlDoc, "input");
            SetOutputAttribute(input, "type", "text", "ContentControl:type");
            SetOutputAttribute(input, "disabled", string.Empty, "ContentControl:disabled");
            if (!string.IsNullOrEmpty(structuredDocumentTag.Text)) {
                SetOutputAttribute(input, "value", structuredDocumentTag.Text!, "ContentControl:value");
            }
            ApplyContentControlMetadata(input, structuredDocumentTag.Alias, structuredDocumentTag.Tag);
            return input;
        }

        static void ReleaseReplacedContentControlContent(
            IDocument htmlDoc,
            OpenXmlElement? sourceControl) {
            if (sourceControl == null) return;
            ReleaseOutputCharacters(htmlDoc, MeasureOutputContentCharacters(sourceControl));
        }

        static bool HasLineBreaks(string? text) =>
            !string.IsNullOrEmpty(text) && (text!.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0);

        static void ApplyContentControlMetadata(IElement element, string? alias, string? tag) {
            if (!string.IsNullOrEmpty(alias)) {
                SetOutputAttribute(element, "aria-label", alias!, "ContentControl:aria-label");
            }

            if (!string.IsNullOrEmpty(tag)) {
                SetOutputAttribute(element, "data-tag", tag!, "ContentControl:data-tag");
            }
        }
    }
}
