using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const string FooterSlotUri = "urn:officeimo:powerpoint:footer-slot:1";
        private bool _footerDefinitionsLoaded;

        /// <summary>Creates an unattached native text body owned by this package, without adding a slide.</summary>
        internal PowerPointLayoutNativeText CreateFooterTextBody() {
            var shape = new Shape(
                new NonVisualShapeProperties(new NonVisualDrawingProperties { Id = 0U, Name = "Footer" },
                    new NonVisualShapeDrawingProperties { TextBox = true }, new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(), new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));
            return new PowerPointLayoutNativeText(new PowerPointTextBox(shape, _presentationPart));
        }

        /// <summary>Changes a deck slot and applies it to existing slides through the normal native footer renderer.</summary>
        internal void SetDeckFooter(OfficeAlignment alignment, PowerPointLayoutFooter footer) {
            EnsureFooterDefinitions();
            LayoutDeck.Footers[alignment] = footer;
            foreach (var slide in _slides) {
                foreach (var box in slide.TextBoxes.Where(box => TryReadFooterSlot(box, out var slot, out bool deck)
                             && deck && slot == alignment).ToArray()) box.Remove();
                var context = new PowerPointLayoutContext(this, PowerPointLayoutTheme.From(slide.AuthoringTheme));
                PowerPointLayoutEngine.RenderNativeFooters(context, new PowerPointLayoutWriter(slide, context.Theme),
                    slide.NativeFooters, slide.SuppressFooterText, slide.SuppressFooterNumber, slide.NativeFooterColor);
            }
        }

        /// <summary>Recovers deck and slide slots from saved DrawingML, independently of editable shape names.</summary>
        internal void EnsureFooterDefinitions() {
            if (_footerDefinitionsLoaded) return;
            _footerDefinitionsLoaded = true;
            if (ReadFooterDefinitions(PresentationRoot, _presentationPart, LayoutDeck.Footers) == null) return;
            foreach (var slide in _slides) {
                var definitions = ReadFooterDefinitions(slide.SlidePart.Slide!, slide.SlidePart, slide.NativeFooters);
                slide.SuppressFooterText = FooterAttribute(definitions, "omitText") == "1";
                slide.SuppressFooterNumber = FooterAttribute(definitions, "omitNumber") == "1";
                string? color = FooterAttribute(definitions, "textColor");
                slide.NativeFooterColor = color != null && OfficeColor.TryParse("#" + color, out var literal) ? literal.ToRgbHex() : null;
            }
        }

        internal static void SetFooterSlot(PowerPointTextBox box, OfficeAlignment alignment, bool deck) {
            var drawing = ((Shape)box.Element).NonVisualShapeProperties!.NonVisualDrawingProperties!;
            var list = drawing.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();
            if (list == null) drawing.Append(list = new A.NonVisualDrawingPropertiesExtensionList());
            foreach (var old in list.Elements<A.NonVisualDrawingPropertiesExtension>().Where(e => e.Uri?.Value == FooterSlotUri).ToArray()) old.Remove();
            var extension = new A.NonVisualDrawingPropertiesExtension { Uri = FooterSlotUri };
            var marker = new OpenXmlUnknownElement("imo", "footer", "urn:officeimo:powerpoint:footer:1");
            marker.SetAttribute(new OpenXmlAttribute("alignment", string.Empty, alignment.ToString()));
            marker.SetAttribute(new OpenXmlAttribute("scope", string.Empty, deck ? "deck" : "slide"));
            extension.Append(marker);
            list.Append(extension);
        }

        internal static bool TryReadFooterSlot(PowerPointTextBox box, out OfficeAlignment alignment, out bool deck) {
            alignment = OfficeAlignment.Left;
            deck = false;
            var marker = FooterMarker(box);
            if (marker == null || !Enum.TryParse(FooterAttribute(marker, "alignment"), out alignment) ||
                alignment is not (OfficeAlignment.Left or OfficeAlignment.Center or OfficeAlignment.Right)) return false;
            string? scope = FooterAttribute(marker, "scope");
            deck = scope == "deck";
            return deck || scope == "slide";
        }
        private static string? FooterAttribute(OpenXmlElement? element, string name) => element?.GetAttributes()
            .FirstOrDefault(attribute => attribute.LocalName == name && string.IsNullOrEmpty(attribute.NamespaceUri)).Value;

        private static OpenXmlElement? FooterMarker(PowerPointTextBox box) => ((Shape)box.Element).NonVisualShapeProperties?.NonVisualDrawingProperties?
                .GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>()?
                .Elements<A.NonVisualDrawingPropertiesExtension>().FirstOrDefault(e => e.Uri?.Value == FooterSlotUri)?.FirstChild;

    }

    public partial class PowerPointSlide {
        internal Dictionary<OfficeAlignment, PowerPointLayoutFooter> NativeFooters { get; } = new();
        internal bool SuppressFooterText { get; set; }
        internal bool SuppressFooterNumber { get; set; }
        internal string? NativeFooterColor { get; set; }
    }
}
