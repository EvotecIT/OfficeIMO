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
