using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTableCell {
        /// <summary>
        ///     Gets the paragraphs contained in the table cell.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> Paragraphs =>
            EnsureTextBody().Elements<A.Paragraph>()
                .Select(paragraph => new PowerPointParagraph(paragraph, _slidePart))
                .ToList();

        /// <summary>
        ///     Adds a paragraph to the table cell.
        /// </summary>
        public PowerPointParagraph AddParagraph(string text = "", Action<PowerPointParagraph>? configure = null, Action<PowerPointTextRun>? run = null) {
            A.TextBody textBody = EnsureTextBody();
            A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            if (templateParagraph != null && IsEmptyPlaceholderParagraph(templateParagraph) && textBody.Elements<A.Paragraph>().Skip(1).FirstOrDefault() == null) {
                templateParagraph.Remove();
            }

            PowerPointParagraph paragraph = AppendParagraph(textBody, text ?? string.Empty, templateParagraph);
            configure?.Invoke(paragraph);
            if (run != null) {
                PowerPointTextRun runWrapper = paragraph.Runs.FirstOrDefault() ?? paragraph.AddRun(text ?? string.Empty);
                run.Invoke(runWrapper);
            }

            return paragraph;
        }

        /// <summary>
        ///     Adds multiple paragraphs to the table cell.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddParagraphs(IEnumerable<string> paragraphs, Action<PowerPointParagraph>? configure = null) {
            if (paragraphs == null) {
                throw new ArgumentNullException(nameof(paragraphs));
            }

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
            if (paragraphs == null) {
                throw new ArgumentNullException(nameof(paragraphs));
            }

            A.TextBody textBody = EnsureTextBody();
            string[] discardedSoundIds = PowerPointEmbeddedSound
                .GetRelationshipIds(textBody);
            A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            textBody.RemoveAllChildren<A.Paragraph>();

            var results = new List<PowerPointParagraph>();
            foreach (string paragraphText in paragraphs) {
                PowerPointParagraph paragraph = AppendParagraph(textBody, paragraphText ?? string.Empty, templateParagraph);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            if (results.Count == 0) {
                textBody.Append(CreateEmptyParagraph(templateParagraph));
            }

            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                discardedSoundIds);

            return results;
        }

        private A.TextBody EnsureTextBody() {
            Cell.TextBody ??= PowerPointTableTextDefaults.CreateTextBody();
            return Cell.TextBody;
        }

        private PowerPointParagraph AppendParagraph(A.TextBody textBody, string text, A.Paragraph? templateParagraph) {
            A.Paragraph paragraph = new();
            if (templateParagraph?.ParagraphProperties != null) {
                paragraph.ParagraphProperties = (A.ParagraphProperties)templateParagraph.ParagraphProperties.CloneNode(true);
            }

            A.Run run = PowerPointTableTextDefaults.CreateRun(text);
            A.RunProperties? templateRunProperties = templateParagraph?
                .Elements<A.Run>()
                .Select(existingRun => existingRun.RunProperties)
                .FirstOrDefault(runProperties => runProperties != null);
            if (templateRunProperties != null) {
                run.RunProperties = CleanRunProperties((A.RunProperties)templateRunProperties.CloneNode(true));
            }

            paragraph.Append(run);

            A.EndParagraphRunProperties? templateEndProperties = templateParagraph?.GetFirstChild<A.EndParagraphRunProperties>();
            if (templateEndProperties != null) {
                paragraph.Append(CleanEndParagraphRunProperties((A.EndParagraphRunProperties)templateEndProperties.CloneNode(true)));
            } else {
                paragraph.Append(new A.EndParagraphRunProperties { Language = PowerPointTableTextDefaults.Language });
            }

            textBody.Append(paragraph);
            return new PowerPointParagraph(paragraph, _slidePart);
        }

        private static A.Paragraph CreateEmptyParagraph(A.Paragraph? templateParagraph) {
            A.Paragraph paragraph = new();
            if (templateParagraph?.ParagraphProperties != null) {
                paragraph.ParagraphProperties = (A.ParagraphProperties)templateParagraph.ParagraphProperties.CloneNode(true);
            }

            A.Run run = PowerPointTableTextDefaults.CreateRun(string.Empty);
            A.RunProperties? templateRunProperties = templateParagraph?
                .Elements<A.Run>()
                .Select(existingRun => existingRun.RunProperties)
                .FirstOrDefault(runProperties => runProperties != null);
            if (templateRunProperties != null) {
                run.RunProperties = CleanRunProperties((A.RunProperties)templateRunProperties.CloneNode(true));
            }

            paragraph.Append(run);

            A.EndParagraphRunProperties? templateEndProperties = templateParagraph?.GetFirstChild<A.EndParagraphRunProperties>();
            if (templateEndProperties != null) {
                paragraph.Append(CleanEndParagraphRunProperties((A.EndParagraphRunProperties)templateEndProperties.CloneNode(true)));
            } else {
                paragraph.Append(new A.EndParagraphRunProperties { Language = PowerPointTableTextDefaults.Language });
            }

            return paragraph;
        }

        private static bool IsEmptyPlaceholderParagraph(A.Paragraph paragraph) {
            foreach (A.Run run in paragraph.Elements<A.Run>()) {
                if (!string.IsNullOrEmpty(run.Text?.Text)) {
                    return false;
                }
            }

            return paragraph.Elements<A.Break>().FirstOrDefault() == null &&
                paragraph.Elements<A.Field>().FirstOrDefault() == null;
        }

        private static A.RunProperties CleanRunProperties(A.RunProperties properties) {
            properties.RemoveAllChildren<A.HyperlinkOnClick>();
            properties.RemoveAllChildren<A.HyperlinkOnMouseOver>();
            properties.Language ??= PowerPointTableTextDefaults.Language;
            return properties;
        }

        private static A.EndParagraphRunProperties CleanEndParagraphRunProperties(A.EndParagraphRunProperties properties) {
            properties.RemoveAllChildren<A.HyperlinkOnClick>();
            properties.RemoveAllChildren<A.HyperlinkOnMouseOver>();
            properties.Language ??= PowerPointTableTextDefaults.Language;
            return properties;
        }
    }
}
