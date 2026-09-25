using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides run enumeration helpers for <see cref="WordParagraph"/>.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Enumerates visible runs in paragraph order, including cached simple-field results and hyperlink runs.
        /// </summary>
        public IEnumerable<WordParagraph> GetRuns() {
            // Defensive: avoid NullReferenceException if a paragraph wrapper was constructed
            // around a missing OpenXml paragraph (shouldn't happen, but keep ingestion robust).
            if (_paragraph == null) yield break;

            foreach (WordParagraph run in EnumerateVisibleRuns(_paragraph, null)) {
                yield return run;
            }
        }

        private IEnumerable<WordParagraph> EnumerateVisibleRuns(
            DocumentFormat.OpenXml.OpenXmlCompositeElement container,
            Hyperlink? hyperlink) {
            foreach (DocumentFormat.OpenXml.OpenXmlElement element in container.ChildElements) {
                if (element is Run runElement) {
                    yield return new WordParagraph(_document, _paragraph!, runElement) { _hyperlink = hyperlink };
                } else if (element is Hyperlink nestedHyperlink) {
                    foreach (WordParagraph run in EnumerateVisibleRuns(nestedHyperlink, nestedHyperlink)) {
                        yield return run;
                    }
                } else if (element is SimpleField field) {
                    foreach (WordParagraph run in EnumerateVisibleRuns(field, hyperlink)) {
                        yield return run;
                    }
                } else if (element is SdtRun sdtRun) {
                    yield return new WordParagraph(_document, _paragraph!, sdtRun);
                }
            }
        }
    }
}

