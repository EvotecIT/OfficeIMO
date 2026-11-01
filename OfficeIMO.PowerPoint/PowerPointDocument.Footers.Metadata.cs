using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const string FooterDefinitionsUri = "urn:officeimo:powerpoint:footer-definitions:1";
        private const string FooterNamespace = "urn:officeimo:powerpoint:footer:1";

        /// <summary>Stores the authored definition separately from its visible, possibly suppressed physical copies.</summary>
        private void PersistFooterDefinitions() {
            EnsureFooterDefinitions();
            if (LayoutDeck.Footers.Count == 0 && !_slides.Any(slide => slide.NativeFooters.Count != 0)) return;
            WriteFooterDefinitions(PresentationRoot, _presentationPart, LayoutDeck.Footers, persistEmpty: true);
            foreach (var slide in _slides) WriteFooterDefinitions(slide.SlidePart.Slide!, slide.SlidePart,
                slide.NativeFooters, slide.SuppressFooterText, slide.SuppressFooterNumber, slide.NativeFooterColor);
        }

        private static void WriteFooterDefinitions(OpenXmlPartRootElement root, OpenXmlPart owner,
            IReadOnlyDictionary<OfficeAlignment, PowerPointLayoutFooter> footers,
            bool omitText = false, bool omitNumber = false, string? color = null, bool persistEmpty = false) {
            var list = root.ChildElements.FirstOrDefault(element => element.LocalName == "extLst");
            var old = list?.ChildElements.FirstOrDefault(element => FooterAttribute(element, "uri") == FooterDefinitionsUri);
            if (!persistEmpty && footers.Count == 0 && !omitText && !omitNumber && color == null) { old?.Remove(); return; }
            if (list == null) {
                list = root is Presentation ? new PresentationExtensionList() : new SlideExtensionList();
                root.Append(list);
            }
            var definitions = new OpenXmlUnknownElement("imo", "footers", FooterNamespace);
            if (omitText) definitions.SetAttribute(new OpenXmlAttribute("omitText", string.Empty, "1"));
            if (omitNumber) definitions.SetAttribute(new OpenXmlAttribute("omitNumber", string.Empty, "1"));
            if (color != null) definitions.SetAttribute(new OpenXmlAttribute("textColor", string.Empty, color));
            foreach (var item in footers.OrderBy(item => item.Key)) {
                var entry = new OpenXmlUnknownElement("imo", "footer", FooterNamespace);
                entry.SetAttribute(new OpenXmlAttribute("alignment", string.Empty, item.Key.ToString()));
                var body = (TextBody)item.Value.Definition.Box.TextBody!.CloneNode(true);
                if (owner is SlidePart slide) PowerPointLayoutNativeText.RehomeRelationships(body,
                    item.Value.Definition.Box.RelationshipOwnerPart, slide);
                entry.Append(body);
                definitions.Append(entry);
            }
            OpenXmlCompositeElement extension = root is Presentation
                ? new PresentationExtension { Uri = FooterDefinitionsUri }
                : new SlideExtension { Uri = FooterDefinitionsUri };
            extension.Append(definitions);
            if (old == null) list.Append(extension);
            else { list.InsertBefore(extension, old); old.Remove(); }
        }

        private static OpenXmlElement? ReadFooterDefinitions(OpenXmlPartRootElement root, OpenXmlPart owner,
            IDictionary<OfficeAlignment, PowerPointLayoutFooter> target) {
            var definitions = root.ChildElements.FirstOrDefault(element => element.LocalName == "extLst")?
                .ChildElements.FirstOrDefault(element => FooterAttribute(element, "uri") == FooterDefinitionsUri)?.FirstChild;
            if (definitions == null || definitions.LocalName != "footers" || definitions.NamespaceUri != FooterNamespace) return null;
            foreach (var entry in definitions.ChildElements) {
                if (entry.LocalName != "footer" || entry.NamespaceUri != FooterNamespace ||
                    !Enum.TryParse(FooterAttribute(entry, "alignment"), out OfficeAlignment alignment) ||
                    alignment is not (OfficeAlignment.Left or OfficeAlignment.Center or OfficeAlignment.Right)) continue;
                var xml = entry.ChildElements.FirstOrDefault(element => element.LocalName == "txBody"
                    && element.NamespaceUri == "http://schemas.openxmlformats.org/presentationml/2006/main");
                if (xml == null) continue;
                var shape = new Shape(new NonVisualShapeProperties(new NonVisualDrawingProperties { Id = 0U, Name = "Footer" },
                    new NonVisualShapeDrawingProperties { TextBox = true }, new ApplicationNonVisualDrawingProperties()),
                    new ShapeProperties(), new TextBody(xml.OuterXml));
                target[alignment] = new PowerPointLayoutFooter(new PowerPointLayoutNativeText(new PowerPointTextBox(shape, owner)));
            }
            return definitions;
        }
    }
}
