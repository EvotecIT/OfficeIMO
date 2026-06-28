using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTextBox {
        /// <summary>
        ///     Adds a bulleted list to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddBullets(IEnumerable<string> bullets, int level = 0,
            char bulletChar = '\u2022', Action<PowerPointParagraph>? configure = null) {
            if (bullets == null) {
                throw new ArgumentNullException(nameof(bullets));
            }

            var results = new List<PowerPointParagraph>();
            foreach (string bullet in bullets) {
                PowerPointParagraph paragraph = AddParagraph(bullet ?? string.Empty);
                paragraph.SetBullet(bulletChar);
                if (level > 0) {
                    paragraph.Level = level;
                }
                PowerPointListParagraphDefaults.Apply(paragraph);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            return results;
        }

        /// <summary>
        ///     Replaces all paragraphs with a bulleted list.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetBullets(IEnumerable<string> bullets, int level = 0,
            char bulletChar = '\u2022', Action<PowerPointParagraph>? configure = null) {
            if (bullets == null) {
                throw new ArgumentNullException(nameof(bullets));
            }

            return ReplaceParagraphs(bullets, (bullet, templateParagraph) => {
                PowerPointParagraph paragraph = AppendParagraph(EnsureTextBody(), bullet ?? string.Empty, templateParagraph);
                paragraph.SetBullet(bulletChar);
                if (level > 0) {
                    paragraph.Level = level;
                }
                PowerPointListParagraphDefaults.Apply(paragraph);
                configure?.Invoke(paragraph);
                return paragraph;
            });
        }

        /// <summary>
        ///     Adds a numbered list to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddNumberedList(IEnumerable<string> items,
            A.TextAutoNumberSchemeValues style, int startAt = 1,
            int level = 0, Action<PowerPointParagraph>? configure = null) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            var results = new List<PowerPointParagraph>();
            bool first = true;
            foreach (string item in items) {
                PowerPointParagraph paragraph = AddParagraph(item ?? string.Empty);
                if (first) {
                    paragraph.SetNumbered(style, startAt);
                    first = false;
                } else {
                    paragraph.SetNumbered(style);
                }
                if (level > 0) {
                    paragraph.Level = level;
                }
                PowerPointListParagraphDefaults.Apply(paragraph);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            return results;
        }

        /// <summary>
        ///     Adds a numbered list to the textbox using the default numbering style.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddNumberedList(IEnumerable<string> items,
            int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return AddNumberedList(items, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        /// <summary>
        ///     Replaces all paragraphs with a numbered list.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetNumberedList(IEnumerable<string> items,
            A.TextAutoNumberSchemeValues style, int startAt = 1,
            int level = 0, Action<PowerPointParagraph>? configure = null) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            bool first = true;
            return ReplaceParagraphs(items, (item, templateParagraph) => {
                PowerPointParagraph paragraph = AppendParagraph(EnsureTextBody(), item ?? string.Empty, templateParagraph);
                if (first) {
                    paragraph.SetNumbered(style, startAt);
                    first = false;
                } else {
                    paragraph.SetNumbered(style);
                }
                if (level > 0) {
                    paragraph.Level = level;
                }
                PowerPointListParagraphDefaults.Apply(paragraph);
                configure?.Invoke(paragraph);
                return paragraph;
            });
        }

        /// <summary>
        ///     Replaces all paragraphs with a numbered list using the default numbering style.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetNumberedList(IEnumerable<string> items,
            int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return SetNumberedList(items, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        /// <summary>
        ///     Adds a new bulleted paragraph to the textbox.
        /// </summary>
        public PowerPointParagraph AddBullet(string text) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetBullet();
            PowerPointListParagraphDefaults.Apply(paragraph);
            return paragraph;
        }

        /// <summary>
        ///     Adds a numbered item to the textbox.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, A.TextAutoNumberSchemeValues style, int startAt = 1) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetNumbered(style, startAt);
            PowerPointListParagraphDefaults.Apply(paragraph);
            return paragraph;
        }

        /// <summary>
        ///     Adds a numbered item to the textbox using the default numbering style.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, int startAt = 1) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetNumbered(startAt);
            PowerPointListParagraphDefaults.Apply(paragraph);
            return paragraph;
        }
    }
}
