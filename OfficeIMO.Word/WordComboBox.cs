using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a combo box content control within a paragraph.
    /// </summary>
    public class WordComboBox : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordComboBox(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        /// <summary>
        /// Gets the display texts of all combo box items.
        /// </summary>
        public IReadOnlyList<string> Items {
            get {
                var combo = _sdtRun.SdtProperties?.Elements<SdtContentComboBox>()?.FirstOrDefault();
                if (combo != null) {
                    return combo.Elements<ListItem>()
                        .Select(li => li.DisplayText?.Value ?? li.Value?.Value ?? string.Empty)
                        .ToList();
                }
                return new List<string>();
            }
        }

        internal IReadOnlyList<(string Value, string DisplayText)> ExportItems {
            get {
                var combo = _sdtRun.SdtProperties?.Elements<SdtContentComboBox>().FirstOrDefault();
                if (combo == null) return Array.Empty<(string Value, string DisplayText)>();
                return WordContentControlListItems.GetExportItems(combo.Elements<ListItem>());
            }
        }

        /// <summary>
        /// Restores distinct HTML datalist values and display labels after the public combo-box
        /// builder has created the Open XML control with its display-text list.
        /// </summary>
        internal void SetImportedItems(
            IReadOnlyList<(string Value, string DisplayText)> items,
            int selectedIndex) {
            var combo = _sdtRun.SdtProperties?.Elements<SdtContentComboBox>().FirstOrDefault()
                ?? throw new InvalidOperationException("Combo box properties are missing from the structured document tag.");
            var selectedItem = WordContentControlListItems.SetImportedItems(
                combo, _sdtRun, items, selectedIndex);
            combo.LastValue = selectedItem?.Value;
        }

        /// <summary>
        /// Gets or sets the currently selected value displayed by the combo box.
        /// </summary>
        public string? SelectedValue {
            get {
                var combo = _sdtRun.SdtProperties?.Elements<SdtContentComboBox>()?.FirstOrDefault();
                var lastValue = combo?.LastValue?.Value;
                if (!string.IsNullOrEmpty(lastValue)) {
                    return lastValue;
                }

                var text = _sdtRun.SdtContentRun?.Descendants<Text>().FirstOrDefault();
                return text?.Text;
            }
            set {
                var combo = _sdtRun.SdtProperties?.Elements<SdtContentComboBox>()?.FirstOrDefault();
                if (combo == null) {
                    throw new InvalidOperationException("Combo box properties are missing from the structured document tag.");
                }

                if (!string.IsNullOrEmpty(value)) {
                    var allowedValues = combo.Elements<ListItem>()
                        .SelectMany(li => new[] { li.Value?.Value, li.DisplayText?.Value })
                        .Where(item => !string.IsNullOrEmpty(item))
                        .ToList();

                    if (!allowedValues.Any(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase))) {
                        throw new ArgumentException("The selected combo box value must match one of the provided items.", nameof(value));
                    }

                    combo.LastValue = value;
                } else {
                    combo.LastValue = null;
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

                text.Text = value ?? string.Empty;
                text.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this combo box control.
        /// </summary>
        public string? Tag {
            get {
                var tag = _sdtRun.SdtProperties?.OfType<Tag>()?.FirstOrDefault();
                return tag?.Val;
            }
            set {
                var properties = _sdtRun.SdtProperties ?? (_sdtRun.SdtProperties = new SdtProperties());
                var tag = properties.OfType<Tag>().FirstOrDefault();
                if (tag == null) {
                    tag = new Tag();
                    properties.Append(tag);
                }
                tag.Val = value;
            }
        }

        /// <summary>
        /// Gets the alias associated with this combo box control.
        /// </summary>
        public string? Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties?.OfType<SdtAlias>()?.FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Removes the combo box from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }
    }
}
