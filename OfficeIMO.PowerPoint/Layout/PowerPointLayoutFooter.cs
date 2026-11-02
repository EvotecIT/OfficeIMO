using System;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

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

        internal bool HasNumber => _repeat.Paragraphs.Any(paragraph => paragraph.Paragraph.Elements<A.Field>()
            .Any(item => item.Type?.Value == "slidenum"));

        internal PowerPointLayoutNativeText TakeBody(bool omitNumber) {
            var body = !_used ? _original : _repeat.CopyForRepeat();
            _used = true;
            if (omitNumber) {
                foreach (var paragraph in body.Paragraphs) {
                    foreach (var field in paragraph.Paragraph.Elements<A.Field>().Where(f => f.Type?.Value == "slidenum").ToArray()) field.Remove();
                    foreach (var run in paragraph.Paragraph.Elements<A.Run>().Where(PowerPointDocument.IsFooterTotal).ToArray()) run.Remove();
                }
            }
            return body;
        }
    }
}
