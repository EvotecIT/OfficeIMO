using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using System.Linq;

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
        public string Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this repeating section control.
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
        /// Removes the repeating section from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }
    }
}
