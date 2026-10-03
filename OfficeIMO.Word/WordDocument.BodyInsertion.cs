using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        // Open XML InsertBefore searches for the previous sibling. Reuse the last
        // appended block only while its current links still prove it is the boundary's predecessor.
        private OpenXmlElement? _lastAppendedBodyBlock;

        internal void AppendBlockToBody(OpenXmlElement element) {
            if (element == null) {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Parent != null) {
                element.Remove();
            }

            var body = BodyRoot;
            var finalSectionProperties = GetFinalSectionPropertiesInsertionBoundary();
            if (finalSectionProperties != null) {
                if (_lastAppendedBodyBlock?.Parent == body
                    && ReferenceEquals(_lastAppendedBodyBlock.NextSibling(), finalSectionProperties)) {
                    body.InsertAfter(element, _lastAppendedBodyBlock);
                } else {
                    body.InsertBefore(element, finalSectionProperties);
                }
            } else {
                body.AppendChild(element);
            }
            _lastAppendedBodyBlock = element;
        }

        internal SectionProperties? GetFinalSectionPropertiesInsertionBoundary() {
            var body = BodyRoot;
            var finalSectionProperties = body.LastChild as SectionProperties;
            if (finalSectionProperties != null) {
                return finalSectionProperties;
            }

            finalSectionProperties = body.Elements<SectionProperties>().LastOrDefault();
            if (finalSectionProperties != null && !ReferenceEquals(body.ChildElements.LastOrDefault(), finalSectionProperties)) {
                finalSectionProperties.Remove();
                body.AppendChild(finalSectionProperties);
            }

            return finalSectionProperties;
        }
    }
}
