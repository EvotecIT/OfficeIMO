using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a picture content control within a paragraph.
    /// </summary>
    public class WordPictureControl : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordPictureControl(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        /// <summary>
        /// Gets the alias associated with this picture control.
        /// </summary>
        public string Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this picture control.
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
        /// Removes the picture control from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }
    }
}
