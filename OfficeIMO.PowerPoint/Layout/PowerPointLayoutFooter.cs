using OfficeIMO.OpenXml.Internal;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

namespace OfficeIMO.PowerPoint {
    /// <summary>Captures authored native footer nodes once, before layout fills in theme defaults.</summary>
    internal sealed class PowerPointLayoutFooter {
        private readonly PowerPointLayoutNativeText _original;
        private readonly PowerPointLayoutNativeText _repeat;
        private bool _used;

        internal PowerPointLayoutFooter(PowerPointLayoutNativeText body, bool alreadyPlaced = false) {
            body.Freeze();
            _original = body;
            _repeat = body.CopyForRepeat();
            _used = alreadyPlaced;
        }

        internal PowerPointLayoutNativeText Definition => _repeat;

        internal bool HasNumber => _repeat.Paragraphs.Any(paragraph => paragraph.Element.Children(Ns.DrawingML, "fld").ToList()
            .Any(PowerPointDocument.IsSlideNumberField));

        internal PowerPointLayoutNativeText TakeBody(bool omitNumber) {
            var body = !_used ? _original : _repeat.CopyForRepeat();
            _used = true;
            if (omitNumber) {
                foreach (var paragraph in body.Paragraphs) {
                    OfficeXmlNode element = paragraph.Element;
                    foreach (OfficeXmlNode field in element.Children(Ns.DrawingML, "fld").ToList().Where(PowerPointDocument.IsSlideNumberField).ToArray()) field.Remove();
                    foreach (OfficeXmlNode run in element.Children(Ns.DrawingML, "r").ToList().Where(PowerPointDocument.IsFooterTotal).ToArray()) run.Remove();
                }
            }
            return body;
        }
    }
}
