using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents notes for a slide.
    /// </summary>
    public partial class PowerPointNotes {
        private const int P = OfficeXmlNamespaces.PresentationML;
        private const int D = OfficeXmlNamespaces.DrawingML;

        private readonly PowerPointSlide _slide;

        internal PowerPointNotes(PowerPointSlide slide) {
            _slide = slide;
        }

        // The slide's notes slide part when it has one, or none.
        private PowerPointPartRef ExistingNotesPart => _slide.PartRef.RelatedPart(PowerPointPartRef.NotesSlideRelationshipType);

        // The notes slide part, created as the notes model first uses it when the slide has none.
        private PowerPointPartRef EnsureNotesPart() =>
            _slide.OwnerDocument?.EnsureNotesSlidePart(_slide) ?? throw new InvalidOperationException("The slide is not part of a presentation.");

        // The p:spTree of a notes slide, or none.
        private static OfficeXmlNode ShapeTreeOf(PowerPointPartRef notes) => notes.Root.Child(P, "cSld").Child(P, "spTree");

        /// <summary>
        ///     Gets or sets the notes text.
        /// </summary>
        public string Text {
            get {
                OfficeXmlNode shape = GetNotesTextShape(ShapeTreeOf(EnsureNotesPart()));
                return ReadParagraphs(shape.Child(P, "txBody"), Environment.NewLine, skipBlank: false);
            }
            set => SetParagraphs(PowerPointTextContent.SplitParagraphs(value));
        }

        /// <summary>Gets the paragraphs in the existing speaker-notes text body without creating a notes part.</summary>
        public IReadOnlyList<PowerPointParagraph> Paragraphs {
            get {
                PowerPointPartRef notes = ExistingNotesPart;
                var paragraphs = new List<PowerPointParagraph>();
                if (notes.IsNone) return paragraphs;
                PowerPointPartRef slide = _slide.PartRef;
                foreach (OfficeXmlNode paragraph in GetNotesTextShape(ShapeTreeOf(notes)).Child(P, "txBody").Children(D, "p")) {
                    paragraphs.Add(new PowerPointParagraph(paragraph, slide, notes, _slide));
                }
                return paragraphs;
            }
        }

        /// <summary>Replaces the speaker notes with paragraphs that can be further styled through their typed runs.</summary>
        public IReadOnlyList<PowerPointParagraph> SetParagraphs(IEnumerable<string> paragraphs) {
            ArgumentNullException.ThrowIfNull(paragraphs);
            // The texts are read before anything changes, so a source that reads the notes sees them as they were.
            List<string> texts = paragraphs as List<string> ?? paragraphs.ToList();

            PowerPointPartRef notes = EnsureNotesPart();
            OfficeXmlNode tree = EnsureShapeTree(notes.Root);
            OfficeXmlNode shape = GetNotesTextShape(tree);
            if (shape.IsNone) shape = tree.Append(CreateNotesPlaceholderShape(tree, NextShapeId(tree)));
            OfficeXmlNode body = shape.Child(P, "txBody");
            if (body.IsNone) {
                body = shape.CreateElement(P, "txBody");
                body.Append(body.CreateElement(D, "bodyPr"));
                body.Append(body.CreateElement(D, "lstStyle"));
                body = shape.InsertOrdered(body, PShapeView.Order);
            }

            OfficeXmlNode template = body.Child(D, "p");
            OfficeXmlNode paragraphProperties = template.Child(D, "pPr");
            paragraphProperties = paragraphProperties.IsNone ? default : paragraphProperties.Clone();
            OfficeXmlNode endProperties = template.Child(D, "endParaRPr");
            endProperties = endProperties.IsNone ? default : endProperties.Clone();
            OfficeXmlNode runProperties = PowerPointTextBox.FirstRunProperties(template);
            runProperties = runProperties.IsNone ? default : runProperties.Clone();

            body.RemoveChildren(D, "p");
            if (texts.Count == 0) texts = new List<string> { string.Empty };
            PowerPointPartRef slide = _slide.PartRef;
            var result = new List<PowerPointParagraph>();
            foreach (string text in texts) {
                OfficeXmlNode paragraph = body.Append(body.CreateElement(D, "p"));
                if (!paragraphProperties.IsNone) paragraph.Append(paragraphProperties.Clone());
                PowerPointTextContent.InsertInlineContent(paragraph, text ?? string.Empty, runProperties, default);
                if (!endProperties.IsNone) paragraph.Append(endProperties.Clone());
                result.Add(new PowerPointParagraph(paragraph, slide, notes, _slide));
            }
            notes.MarkEdited();
            return result;
        }

        /// <summary>
        ///     Reads existing speaker notes without creating notes or notes-master parts when none exist.
        /// </summary>
        public bool TryGetText(out string text) {
            PowerPointPartRef notes = ExistingNotesPart;
            OfficeXmlNode body = notes.IsNone ? default : GetNotesTextShape(ShapeTreeOf(notes)).Child(P, "txBody");
            if (body.IsNone) {
                text = string.Empty;
                return false;
            }

            text = ReadParagraphs(body, Environment.NewLine, skipBlank: false);
            return text.Length > 0;
        }

        /// <summary>
        /// Tries to read speaker notes that already exist without creating a notes part.
        /// </summary>
        /// <param name="text">The existing speaker-notes text, or an empty string when none exists.</param>
        /// <returns><see langword="true"/> when non-empty speaker notes were found.</returns>
        public bool TryGetExistingText(out string text) {
            text = string.Empty;
            try {
                PowerPointPartRef notes = ExistingNotesPart;
                OfficeXmlNode tree = notes.IsNone ? default : ShapeTreeOf(notes);
                if (tree.IsNone) {
                    return false;
                }

                var blocks = new List<string>();
                foreach (OfficeXmlNode shape in tree.Children(P, "sp")) {
                    string block = ReadParagraphs(shape.Child(P, "txBody"), Environment.NewLine, skipBlank: true);
                    if (!string.IsNullOrWhiteSpace(block)) blocks.Add(block);
                }
                text = string.Join(Environment.NewLine + Environment.NewLine, blocks).Trim();
                return text.Length > 0;
            } catch {
                text = string.Empty;
                return false;
            }
        }

        /// <summary>Writes the notes slide of a slide the SDK holds; notes on the package are written with it.</summary>
        internal void Save() {
            if (_slide.OwnerDocument is { IsSdkLoaded: true }) PowerPointUtils.SaveNotesSlide(_slide.SdkSlidePart);
        }

        // The notes slide's shape tree, with the group properties the model adds when they are missing.
        private static OfficeXmlNode EnsureShapeTree(OfficeXmlNode root) {
            OfficeXmlNode data = root.GetOrAddChild(P, "cSld", PNotesSlideView.Order);
            OfficeXmlNode tree = data.GetOrAddChild(P, "spTree", PCommonSlideDataView.Order);
            if (tree.Child(P, "nvGrpSpPr").IsNone) {
                OfficeXmlNode group = tree.CreateElement(P, "nvGrpSpPr");
                OfficeXmlNode drawing = group.Append(group.CreateElement(P, "cNvPr"));
                drawing.SetAttribute("id", "1");
                drawing.SetAttribute("name", "Notes Group Shape");
                group.Append(group.CreateElement(P, "cNvGrpSpPr"));
                group.Append(group.CreateElement(P, "nvPr"));
                tree.InsertBefore(group, tree.FirstChild);
            }
            if (tree.Child(P, "grpSpPr").IsNone) tree.Append(CreateGroupShapeProperties(tree));
            return tree;
        }

        // The group properties a new shape tree gets: a transform with zero offset and extents for itself and its children.
        private static OfficeXmlNode CreateGroupShapeProperties(OfficeXmlNode context) {
            OfficeXmlNode properties = context.CreateElement(P, "grpSpPr");
            OfficeXmlNode transform = properties.Append(properties.CreateElement(D, "xfrm"));
            foreach ((string name, string first, string second) in new[] { ("off", "x", "y"), ("ext", "cx", "cy"), ("chOff", "x", "y"), ("chExt", "cx", "cy") }) {
                OfficeXmlNode point = transform.Append(transform.CreateElement(D, name));
                point.SetAttribute(first, "0");
                point.SetAttribute(second, "0");
            }
            return properties;
        }

        // The body placeholder the model adds to notes without one: an empty run and end properties in en-US.
        private static OfficeXmlNode CreateNotesPlaceholderShape(OfficeXmlNode context, uint id) {
            OfficeXmlNode shape = context.CreateElement(P, "sp");
            OfficeXmlNode nonVisual = shape.Append(shape.CreateElement(P, "nvSpPr"));
            OfficeXmlNode drawing = nonVisual.Append(nonVisual.CreateElement(P, "cNvPr"));
            drawing.SetAttribute("id", id.ToString(System.Globalization.CultureInfo.InvariantCulture));
            drawing.SetAttribute("name", "Notes Placeholder 1");
            nonVisual.Append(nonVisual.CreateElement(P, "cNvSpPr")).Append(nonVisual.CreateElement(D, "spLocks")).SetAttribute("noGrp", "1");
            OfficeXmlNode placeholder = nonVisual.Append(nonVisual.CreateElement(P, "nvPr")).Append(nonVisual.CreateElement(P, "ph"));
            placeholder.SetAttribute("type", "body");
            placeholder.SetAttribute("idx", "1");
            shape.Append(shape.CreateElement(P, "spPr"));
            OfficeXmlNode body = shape.Append(shape.CreateElement(P, "txBody"));
            body.Append(body.CreateElement(D, "bodyPr"));
            body.Append(body.CreateElement(D, "lstStyle"));
            body.Append(PowerPointTableTextDefaults.CreateParagraph(body));
            return shape;
        }

        // One more than the largest drawing id in the shape tree.
        private static uint NextShapeId(OfficeXmlNode tree) {
            uint max = 0U;
            foreach (OfficeXmlNode drawing in tree.Descendants(P, "cNvPr")) {
                if (drawing.UInt32Attribute("id") is { } id && id > max) max = id;
            }
            return max + 1U;
        }

        // The body placeholder, or else the first shape with text.
        private static OfficeXmlNode GetNotesTextShape(OfficeXmlNode tree) {
            OfficeXmlNode withText = default;
            foreach (OfficeXmlNode shape in tree.Children(P, "sp")) {
                if (shape.Child(P, "nvSpPr").Child(P, "nvPr").Child(P, "ph").Attribute("type") == "body") return shape;
                if (withText.IsNone && !shape.Child(P, "txBody").IsNone) withText = shape;
            }
            return withText;
        }

        // The text of a body's paragraphs joined by a separator, leaving out blank ones when asked.
        private static string ReadParagraphs(OfficeXmlNode body, string separator, bool skipBlank) {
            var paragraphs = new List<string>();
            foreach (OfficeXmlNode paragraph in body.Children(D, "p")) {
                string text = ReadParagraphText(paragraph);
                if (!skipBlank || !string.IsNullOrWhiteSpace(text)) paragraphs.Add(text);
            }
            return string.Join(separator, paragraphs);
        }

        /// <summary>The text of a notes paragraph: runs and fields, with a new line for each break.</summary>
        internal static string ReadParagraphText(OfficeXmlNode paragraph) {
            var builder = new StringBuilder();
            foreach (OfficeXmlNode child in paragraph.Children()) {
                if (child.NamespaceId != D) continue;
                switch (child.LocalName) {
                    case "r":
                    case "fld":
                        builder.Append(child.ChildText(D, "t") ?? string.Empty);
                        break;
                    case "br":
                        builder.AppendLine();
                        break;
                }
            }
            return builder.ToString();
        }
    }
}
