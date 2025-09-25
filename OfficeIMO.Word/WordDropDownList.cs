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
}
