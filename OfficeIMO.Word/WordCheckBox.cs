using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a checkbox content control within a paragraph.
    /// </summary>
    public class WordCheckBox : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordCheckBox(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        /// <summary>
        /// Gets or sets whether the checkbox is checked.
        /// </summary>
        public bool IsChecked {
            get {
                var cb = _sdtRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().FirstOrDefault();
                var ch = cb?.Elements<W14.Checked>().FirstOrDefault();
                return ch != null && ch.Val != null && ch.Val.Value == W14.OnOffValues.One;
            }
            set {
                var cb = _sdtRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().FirstOrDefault();
                if (cb != null) {
                    var ch = cb.Elements<W14.Checked>().FirstOrDefault();
                    if (ch == null) {
                        ch = new W14.Checked();
                        cb.Append(ch);
                    }
                    ch.Val = value ? W14.OnOffValues.One : W14.OnOffValues.Zero;
                }
            }
        }

        /// <summary>
        /// Gets the alias associated with this checkbox control.
        /// </summary>
        public string? Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties?.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this checkbox control.
        /// </summary>
        public string? Tag {
            get {
                var tag = _sdtRun.SdtProperties?.OfType<Tag>().FirstOrDefault();
                return tag?.Val;
            }
            set {
                var tag = _sdtRun.SdtProperties?.OfType<Tag>().FirstOrDefault();
                if (tag == null) {
                    tag = new Tag();
                    _sdtRun.SdtProperties?.Append(tag);
                }
                if (tag != null) {
                    tag.Val = value;
                }
            }
        }

        /// <summary>
        /// Removes the checkbox from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }
    }
}
