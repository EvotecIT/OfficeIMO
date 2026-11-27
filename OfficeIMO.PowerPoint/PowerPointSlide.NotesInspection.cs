using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        // The slide's notes slide (p:notes), or none.
        private OfficeXmlNode NotesRoot {
            get {
                OfficeXmlNode root = PartRef.RelatedPart(PowerPointPartRef.NotesSlideRelationshipType).Root;
                return root.Is(OfficeXmlNamespaces.PresentationML, "notes") ? root : default;
            }
        }

        /// <summary>Whether the slide already contains a speaker-notes part.</summary>
        public bool HasSpeakerNotes => !NotesRoot.IsNone;

        /// <summary>Reads existing speaker-note text without creating a notes part as a side effect.</summary>
        public string GetSpeakerNotesText() {
            OfficeXmlNode tree = NotesRoot.Child(OfficeXmlNamespaces.PresentationML, "cSld").Child(OfficeXmlNamespaces.PresentationML, "spTree");
            var blocks = new List<string>();
            foreach (OfficeXmlNode shape in tree.Children(OfficeXmlNamespaces.PresentationML, "sp")) {
                var paragraphs = new List<string>();
                foreach (OfficeXmlNode paragraph in shape.Child(OfficeXmlNamespaces.PresentationML, "txBody").Children(OfficeXmlNamespaces.DrawingML, "p")) {
                    string text = PowerPointNotes.ReadParagraphText(paragraph).Trim();
                    if (!string.IsNullOrWhiteSpace(text)) paragraphs.Add(text);
                }
                string block = string.Join("\n", paragraphs);
                if (!string.IsNullOrWhiteSpace(block)) blocks.Add(block);
            }
            return string.Join("\n\n", blocks);
        }
    }
}
