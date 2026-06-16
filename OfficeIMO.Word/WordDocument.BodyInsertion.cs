using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        internal void AppendBlockToBody(OpenXmlElement element) {
            if (element == null) {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Parent != null) {
                element.Remove();
            }

            var body = BodyRoot;
            var finalSectionProperties = body.Elements<SectionProperties>().LastOrDefault();
            if (finalSectionProperties != null) {
                body.InsertBefore(element, finalSectionProperties);
            } else {
                body.AppendChild(element);
            }
        }
    }
}
