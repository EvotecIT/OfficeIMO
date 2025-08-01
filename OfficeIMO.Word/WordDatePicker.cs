using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.Xml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a date picker content control within a paragraph.
    /// </summary>
    public class WordDatePicker : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordDatePicker(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
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
                var dp = _sdtRun.SdtProperties.Elements<SdtContentDate>().FirstOrDefault();
                if (dp == null) {
                    dp = new SdtContentDate();
                    _sdtRun.SdtProperties.Append(dp);
                }
                dp.FullDate = value.HasValue ? new DateTimeValue(value.Value) : null;
            }
        }

        /// <summary>
        /// Gets the alias associated with this date picker control.
        /// </summary>
        public string Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this date picker control.
        /// </summary>
        public string Tag {
            get {
                var tag = _sdtRun.SdtProperties.OfType<Tag>().FirstOrDefault();
                return tag?.Val;
            }
            set {
                var tag = _sdtRun.SdtProperties.OfType<Tag>().FirstOrDefault();
                if (tag == null) {
                    tag = new Tag();
                    _sdtRun.SdtProperties.Append(tag);
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
    }
}
