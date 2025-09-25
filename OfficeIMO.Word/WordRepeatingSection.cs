using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a repeating section content control within a paragraph.
    /// </summary>
    public class WordRepeatingSection : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordRepeatingSection(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        /// <summary>
        /// Gets the alias associated with this repeating section control.
        /// </summary>
        public string? Alias {
            get {
                var properties = _sdtRun.SdtProperties;
                var sdtAlias = properties?.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this repeating section control.
        /// </summary>
        public string? Tag {
            get {
                var properties = _sdtRun.SdtProperties;
                var tag = properties?.OfType<Tag>().FirstOrDefault();
                return tag?.Val;
            }
            set {
                var properties = EnsureProperties();
                var tag = properties.OfType<Tag>().FirstOrDefault();
                if (value == null) {
                    tag?.Remove();
                    return;
                }
                if (tag == null) {
                    tag = new Tag();
                    properties.Append(tag);
                }
                tag.Val = value;
            }
        }

        /// <summary>
        /// Removes the repeating section from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }

        private SdtProperties EnsureProperties() {
            if (_sdtRun.SdtProperties == null) {
                _sdtRun.SdtProperties = new SdtProperties();
            }
            return _sdtRun.SdtProperties;
        }
    }
}
