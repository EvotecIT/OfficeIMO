using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a dropdown list content control within a paragraph.
    /// </summary>
    public class WordDropDownList : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordDropDownList(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        /// <summary>
        /// Gets the display texts of all list items.
        /// </summary>
        public IReadOnlyList<string> Items {
            get {
                var ddl = _sdtRun.SdtProperties?.Elements<SdtContentDropDownList>().FirstOrDefault();
                if (ddl != null) {
                    return ddl.Elements<ListItem>().Select(li => li.DisplayText?.Value ?? li.Value?.Value).ToList();
                }
                return new List<string>();
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this dropdown list control.
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
        /// Gets the alias associated with this dropdown list control.
        /// </summary>
        public string Alias {
            get {
                var sdtAlias = _sdtRun.SdtProperties.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Removes the dropdown list from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }
    }
}
