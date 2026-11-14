using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        // Open XML InsertBefore searches for the previous sibling. Reuse the last
        // appended block only while its current links still prove it is the boundary's predecessor.
        private OpenXmlElement? _lastAppendedBodyBlock;

        internal void AppendBlockToBody(OpenXmlElement element) {
            ArgumentNullException.ThrowIfNull(element);

            if (element.Parent != null) {
                element.Remove();
            }

            var body = BodyRoot;
            var finalSectionProperties = GetFinalSectionPropertiesInsertionBoundary(body);
            if (finalSectionProperties != null) {
                InsertBodyBlockBefore(body, element, finalSectionProperties);
            } else {
                body.AppendChild(element);
                _lastAppendedBodyBlock = element;
            }
        }

        /// <summary>Inserts a body block before <paramref name="boundary"/>, a direct child of the body.</summary>
        internal void InsertBodyBlockBefore(Body body, OpenXmlElement element, OpenXmlElement boundary) {
            if (_lastAppendedBodyBlock?.Parent == body
                && ReferenceEquals(_lastAppendedBodyBlock.NextSibling(), boundary)) {
                body.InsertAfter(element, _lastAppendedBodyBlock);
            } else {
                body.InsertBefore(element, boundary);
            }
            _lastAppendedBodyBlock = element;
        }

        internal SectionProperties? GetFinalSectionPropertiesInsertionBoundary() => GetFinalSectionPropertiesInsertionBoundary(BodyRoot);

        private static SectionProperties? GetFinalSectionPropertiesInsertionBoundary(Body body) {
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
