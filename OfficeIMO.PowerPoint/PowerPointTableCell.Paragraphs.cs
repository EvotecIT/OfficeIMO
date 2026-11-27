using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTableCell {
        /// <summary>
        ///     Gets the paragraphs contained in the table cell.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> Paragraphs {
            get {
                var paragraphs = new List<PowerPointParagraph>();
                foreach (OfficeXmlNode paragraph in EnsureTextBody().Children(D, "p")) paragraphs.Add(ParagraphFor(paragraph));
                return paragraphs;
            }
        }

        /// <summary>
        ///     Adds a paragraph to the table cell.
        /// </summary>
        public PowerPointParagraph AddParagraph(string text = "", Action<PowerPointParagraph>? configure = null, Action<PowerPointTextRun>? run = null) {
            OfficeXmlNode body = EnsureTextBody();
            OfficeXmlNode template = body.Child(D, "p");
            // The empty paragraph a new cell starts with gives way to the first one added.
            if (!template.IsNone && IsEmptyPlaceholderParagraph(template) && template.NextSiblingNamed(D, "p").IsNone) {
                template.Remove();
                MarkEdited();
            }

            PowerPointParagraph paragraph = AppendParagraph(body, text ?? string.Empty, template, out PowerPointTextRun authoredRun);
            configure?.Invoke(paragraph);
            if (run != null) run(authoredRun.BelongsTo(paragraph) ? authoredRun : paragraph.AddText(text ?? string.Empty));

            return paragraph;
        }

        /// <summary>
        ///     Adds multiple paragraphs to the table cell.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddParagraphs(IEnumerable<string> paragraphs, Action<PowerPointParagraph>? configure = null) {
            ArgumentNullException.ThrowIfNull(paragraphs);

            var results = new List<PowerPointParagraph>();
            foreach (string paragraphText in paragraphs) {
                PowerPointParagraph paragraph = AddParagraph(paragraphText ?? string.Empty);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            return results;
        }

        /// <summary>
        ///     Replaces all table-cell paragraphs with the provided content.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetParagraphs(IEnumerable<string> paragraphs, Action<PowerPointParagraph>? configure = null) {
            ArgumentNullException.ThrowIfNull(paragraphs);
            return ReplaceTableCellParagraphs(paragraphs, (paragraphText, template) => {
                PowerPointParagraph paragraph = AppendParagraph(EnsureTextBody(), paragraphText ?? string.Empty, template);
                configure?.Invoke(paragraph);
                return paragraph;
            });
        }

        // Replaces the paragraphs with those the callback adds, each copying from the first paragraph the cell had; with
        // none added, one empty paragraph like it stays.
        private IReadOnlyList<PowerPointParagraph> ReplaceTableCellParagraphs<T>(IEnumerable<T> items, Func<T, OfficeXmlNode, PowerPointParagraph> addParagraph) {
            // The items are read before anything is removed, so a source that changes the cell sees it unchanged.
            List<T> list = items as List<T> ?? items.ToList();
            OfficeXmlNode body = EnsureTextBody();
            List<string> soundIds = PowerPointTextBox.SoundIds(body);
            OfficeXmlNode template = body.Child(D, "p");
            if (!template.IsNone) template.Remove();
            body.RemoveChildren(D, "p");
            MarkEdited();

            var results = new List<PowerPointParagraph>();
            foreach (T item in list) {
                results.Add(addParagraph(item, template));
            }

            if (results.Count == 0) {
                body.Append(CreateEmptyParagraph(body, template));
            }

            RemoveUnusedSounds(soundIds);
            return results;
        }

        private void RemoveUnusedSounds(List<string> soundIds) {
            if (soundIds.Count == 0) return;
            PowerPointPartRef owner = OwnerPartRef;
            if (owner.IsNone) return;
            foreach (string id in soundIds) owner.RemoveRelationshipIfUnused(id, default);
        }

        // The cell's text body; a cell without one gets the default body with one empty paragraph.
        private OfficeXmlNode EnsureTextBody() {
            OfficeXmlNode body = BodyElement;
            return body.IsNone ? CreateBody(withParagraph: true) : body;
        }

        private PowerPointParagraph AppendParagraph(OfficeXmlNode body, string text, OfficeXmlNode template) =>
            AppendParagraph(body, text, template, out _);

        // Appends a paragraph with the template's paragraph and end properties and one run of text with its first run's
        // formatting, without hyperlinks and with a language.
        private PowerPointParagraph AppendParagraph(OfficeXmlNode body, string text, OfficeXmlNode template, out PowerPointTextRun authoredRun) {
            OfficeXmlNode paragraph = body.CreateElement(D, "p");
            OfficeXmlNode paragraphProperties = template.Child(D, "pPr");
            if (!paragraphProperties.IsNone) paragraph.Append(paragraphProperties.Clone());
            paragraph.Append(CleanProperties(template.Child(D, "endParaRPr"), paragraph, "endParaRPr"));
            OfficeXmlNode runProperties = PowerPointTextBox.FirstRunProperties(template);

            paragraph = body.Append(paragraph);
            MarkEdited();
            PowerPointParagraph wrapper = ParagraphFor(paragraph);
            authoredRun = wrapper.InsertRun(text, CleanProperties(runProperties, paragraph, "rPr"));
            return wrapper;
        }

        // A detached empty paragraph after the template: its properties, one empty run and its end properties.
        private static OfficeXmlNode CreateEmptyParagraph(OfficeXmlNode context, OfficeXmlNode template) {
            OfficeXmlNode paragraph = context.CreateElement(D, "p");
            OfficeXmlNode paragraphProperties = template.Child(D, "pPr");
            if (!paragraphProperties.IsNone) paragraph.Append(paragraphProperties.Clone());
            OfficeXmlNode run = paragraph.Append(paragraph.CreateElement(D, "r"));
            OfficeXmlNode runProperties = PowerPointTextBox.FirstRunProperties(template);
            run.Append(runProperties.IsNone ? PowerPointTableTextDefaults.CreateProperties(run, "rPr") : CleanProperties(runProperties, run, "rPr"));
            run.Append(run.CreateElement(D, "t"));
            paragraph.Append(CleanProperties(template.Child(D, "endParaRPr"), paragraph, "endParaRPr"));
            return paragraph;
        }

        private static bool IsEmptyPlaceholderParagraph(OfficeXmlNode paragraph) {
            foreach (OfficeXmlNode run in paragraph.Children(D, "r")) {
                if (!string.IsNullOrEmpty(run.ChildText(D, "t"))) return false;
            }
            return paragraph.Child(D, "br").IsNone && paragraph.Child(D, "fld").IsNone;
        }

        // A copy of run or end properties without hyperlinks and with a language; new properties when there are none.
        private static OfficeXmlNode CleanProperties(OfficeXmlNode properties, OfficeXmlNode context, string localName) {
            if (properties.IsNone) return PowerPointTableTextDefaults.CreateProperties(context, localName);
            OfficeXmlNode copy = properties.Clone();
            WithoutHyperlinks(copy);
            if (copy.Attribute("lang") == null) copy.SetAttribute("lang", PowerPointTableTextDefaults.Language);
            return copy;
        }
    }
}
