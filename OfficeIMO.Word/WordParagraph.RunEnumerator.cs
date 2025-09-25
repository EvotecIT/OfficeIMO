using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides run enumeration helpers for <see cref="WordParagraph"/>.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Enumerates all runs within the paragraph, including runs nested in hyperlinks.
        /// </summary>
        public IEnumerable<WordParagraph> GetRuns() {
            foreach (var element in _paragraph.ChildElements) {
                if (element is Run runElement) {
                    yield return new WordParagraph(_document, _paragraph, runElement);
                } else if (element is Hyperlink hyperlink) {
                    foreach (var childRun in hyperlink.Elements<Run>()) {
                        var paragraph = new WordParagraph(_document, _paragraph, childRun) { _hyperlink = hyperlink };
                        yield return paragraph;
                    }
                }
            }
        }
    }
}

