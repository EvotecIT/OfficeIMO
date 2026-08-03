using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a dropdown list content control within a paragraph.
    /// </summary>
    public class WordDropDownList : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordDropDownList(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            _sdtRun = sdtRun ?? throw new ArgumentNullException(nameof(sdtRun));
        }

        /// <summary>
        /// Gets the display texts of all list items.
        /// </summary>
        public IReadOnlyList<string> Items {
            get {
                var ddl = _sdtRun.SdtProperties?.Elements<SdtContentDropDownList>().FirstOrDefault();
                if (ddl != null) {
                    return ddl.Elements<ListItem>()
                        .Select(li => li.DisplayText?.Value ?? li.Value?.Value ?? string.Empty)
                        .ToList();
                }
                return Array.Empty<string>();
            }
        }

        internal IReadOnlyList<(string Value, string DisplayText)> ExportItems {
            get {
                var ddl = _sdtRun.SdtProperties?.Elements<SdtContentDropDownList>().FirstOrDefault();
                if (ddl == null) return Array.Empty<(string Value, string DisplayText)>();
                return WordContentControlListItems.GetExportItems(ddl.Elements<ListItem>());
            }
        }

        /// <summary>
        /// Restores distinct HTML option values and display labels after the public dropdown
        /// builder has created the Open XML control with its display-text list.
        /// </summary>
        internal void SetImportedItems(
            IReadOnlyList<(string Value, string DisplayText)> items,
            int selectedIndex) {
            var dropDown = _sdtRun.SdtProperties?.Elements<SdtContentDropDownList>().FirstOrDefault()
                ?? throw new InvalidOperationException("Dropdown list properties are missing from the structured document tag.");
            var selectedItem = WordContentControlListItems.SetImportedItems(
                dropDown, _sdtRun, items, selectedIndex);
            dropDown.LastValue = selectedItem.Value;
        }

        /// <summary>
        /// Gets the selected item's internal value, or selects an item by display text or internal value.
        /// </summary>
        public string? SelectedValue {
            get {
                var dropDown = _sdtRun.SdtProperties?.Elements<SdtContentDropDownList>()
                    .FirstOrDefault();
                if (dropDown?.LastValue != null) {
                    return dropDown.LastValue.Value ?? string.Empty;
                }

                var text = _sdtRun.SdtContentRun?.Descendants<Text>().FirstOrDefault();
                return text?.Text;
            }
            set {
                var ddl = _sdtRun.SdtProperties?.Elements<SdtContentDropDownList>().FirstOrDefault();
                if (ddl == null) {
                    throw new InvalidOperationException("Dropdown list properties are missing from the structured document tag.");
                }

                ListItem? selectedItem = null;
                if (!string.IsNullOrEmpty(value)) {
                    List<ListItem> items = ddl.Elements<ListItem>().ToList();
                    selectedItem = items.FirstOrDefault(item =>
                        string.Equals(item.DisplayText?.Value, value, StringComparison.OrdinalIgnoreCase))
                        ?? items.FirstOrDefault(item =>
                            string.Equals(item.Value?.Value, value, StringComparison.OrdinalIgnoreCase));

                    if (selectedItem == null) {
                        throw new ArgumentException("The selected dropdown list value must match one of the provided items.", nameof(value));
                    }
                    ddl.LastValue = selectedItem.Value?.Value
                        ?? selectedItem.DisplayText?.Value;
                } else {
                    ddl.LastValue = null;
                }

                var content = _sdtRun.SdtContentRun ?? (_sdtRun.SdtContentRun = new SdtContentRun());
                var run = content.Elements<Run>().FirstOrDefault();
                if (run == null) {
                    run = new Run();
                    content.Append(run);
                }

                var text = run.Elements<Text>().FirstOrDefault();
                if (text == null) {
                    text = new Text();
                    run.Append(text);
                }

                text.Text = selectedItem?.DisplayText?.Value
                    ?? selectedItem?.Value?.Value
                    ?? string.Empty;
                text.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        internal bool TryGetSelectedInternalValue(out string value) {
            var dropDown = _sdtRun.SdtProperties?.Elements<SdtContentDropDownList>()
                .FirstOrDefault();
            if (dropDown?.LastValue != null) {
                value = dropDown.LastValue.Value ?? string.Empty;
                return true;
            }
            value = string.Empty;
            return false;
        }

        /// <summary>
        /// Gets or sets the tag value for this dropdown list control.
        /// </summary>
        public string? Tag {
            get {
                var tag = _sdtRun.SdtProperties?.OfType<Tag>().FirstOrDefault();
                return tag?.Val;
            }
            set {
                var properties = EnsureProperties();
                var tag = properties.OfType<Tag>().FirstOrDefault();
                if (tag == null) {
                    tag = new Tag();
                    properties.Append(tag);
                }
                tag.Val = value;
            }
        }

        /// <summary>
        /// Gets the alias associated with this dropdown list control.
        /// </summary>
        public string? Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties?.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Removes the dropdown list from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }

        private SdtProperties EnsureProperties() {
            var properties = _sdtRun.SdtProperties;
            if (properties == null) {
                properties = new SdtProperties();
                _sdtRun.SdtProperties = properties;
            }

            return properties;
        }
    }

    internal static class WordContentControlListItems {
        internal static IReadOnlyList<(string Value, string DisplayText)> GetExportItems(
            IEnumerable<ListItem> items) => items
                .Select(item => (
                    item.Value?.Value ?? item.DisplayText?.Value ?? string.Empty,
                    item.DisplayText?.Value ?? item.Value?.Value ?? string.Empty))
                .ToList();

        internal static (string Value, string DisplayText) SetImportedItems(
            OpenXmlCompositeElement listContainer,
            SdtRun sdtRun,
            IReadOnlyList<(string Value, string DisplayText)> items,
            int selectedIndex) {
            listContainer.RemoveAllChildren<ListItem>();
            foreach ((string value, string displayText) in items) {
                listContainer.Append(new ListItem { Value = value, DisplayText = displayText });
            }

            (string selectedValue, string selectedDisplayText) = items[selectedIndex];
            var content = sdtRun.SdtContentRun ?? (sdtRun.SdtContentRun = new SdtContentRun());
            var run = content.Elements<Run>().FirstOrDefault();
            if (run == null) {
                run = new Run();
                content.Append(run);
            }
            var text = run.Elements<Text>().FirstOrDefault();
            if (text == null) {
                text = new Text();
                run.Append(text);
            }
            text.Text = selectedDisplayText;
            text.Space = SpaceProcessingModeValues.Preserve;
            return (selectedValue, selectedDisplayText);
        }
    }
}
