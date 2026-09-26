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

            var fieldVisibility = WordComplexFieldRunVisibility.ForParagraph(_paragraph);
            foreach (WordParagraph run in EnumerateVisibleRuns(_paragraph, null, fieldVisibility)) {
                yield return run;
            }
        }

        private IEnumerable<WordParagraph> EnumerateVisibleRuns(
            DocumentFormat.OpenXml.OpenXmlCompositeElement container,
            Hyperlink? hyperlink,
            WordComplexFieldRunVisibility fieldVisibility) {
            foreach (DocumentFormat.OpenXml.OpenXmlElement element in container.ChildElements) {
                if (element is Run runElement) {
                    Run? visibleRun = fieldVisibility.GetVisibleRun(runElement, out var visibleSourceChildren);
                    if (visibleRun != null)
                        yield return new WordParagraph(_document, _paragraph!, runElement) {
                            _hyperlink = hyperlink,
                            _visibleRun = ReferenceEquals(visibleRun, runElement) ? null : visibleRun,
                            _visibleRunSourceChildren = visibleSourceChildren
                        };
                } else if (element is CustomXmlRun customXml) {
                    foreach (WordParagraph run in EnumerateVisibleRuns(customXml, hyperlink, fieldVisibility)) {
                        yield return run;
                    }
                } else if (element is Hyperlink nestedHyperlink) {
                    foreach (WordParagraph run in EnumerateVisibleRuns(nestedHyperlink, nestedHyperlink, fieldVisibility)) {
                        yield return run;
                    }
                } else if (element is SimpleField field) {
                    if (fieldVisibility.IsVisible) {
                        foreach (WordParagraph run in EnumerateVisibleRuns(field, hyperlink, fieldVisibility)) {
                            yield return run;
                        }
                    } else {
                        fieldVisibility.ObserveDescendantRuns(field);
                    }
                } else if (element is SdtRun sdtRun) {
                    if (sdtRun.Descendants<FieldChar>().Any()) {
                        if (sdtRun.SdtContentRun != null) {
                            foreach (WordParagraph run in EnumerateVisibleRuns(sdtRun.SdtContentRun, hyperlink, fieldVisibility))
                                yield return run;
                        }
                    } else if (fieldVisibility.IsVisible) {
                        yield return new WordParagraph(_document, _paragraph!, sdtRun) { _hyperlink = hyperlink };
                    }
                }
            }
        }
    }
}

