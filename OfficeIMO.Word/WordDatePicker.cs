using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a date picker content control within a paragraph.
    /// </summary>
    public class WordDatePicker : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordDatePicker(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            _sdtRun = sdtRun ?? throw new ArgumentNullException(nameof(sdtRun));
        }

        /// <summary>
        /// Gets or sets the selected date.
        /// </summary>
        public DateTime? Date {
            get {
                var dp = _sdtRun.SdtProperties?.Elements<SdtContentDate>().FirstOrDefault();
                if (dp?.FullDate != null) {
                    return dp.FullDate.Value;
                }
                return null;
            }
            set {
                var properties = EnsureProperties();
                var dp = properties.Elements<SdtContentDate>().FirstOrDefault();
                if (dp == null) {
                    dp = new SdtContentDate();
                    properties.Append(dp);
                }
                dp.FullDate = value.HasValue ? new DateTimeValue(value.Value) : null;
            }
        }

        /// <summary>
        /// Gets the alias associated with this date picker control.
        /// </summary>
        public string? Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties?.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this date picker control.
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
        /// Removes the date picker from the paragraph.
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
