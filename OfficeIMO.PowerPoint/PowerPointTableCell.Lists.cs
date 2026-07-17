using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTableCell {
        /// <summary>
        ///     Adds a bulleted paragraph to the table cell.
        /// </summary>
        public PowerPointParagraph AddBullet(string text, char bulletChar = '\u2022', int level = 0, Action<PowerPointParagraph>? configure = null) {
            PowerPointParagraph paragraph = AddParagraph(text ?? string.Empty);
            paragraph.SetBullet(bulletChar);
            if (level > 0) {
                paragraph.Level = level;
            }

            PowerPointListParagraphDefaults.Apply(paragraph);
            configure?.Invoke(paragraph);
            return paragraph;
        }

        /// <summary>
        ///     Adds a bulleted list to the table cell.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddBullets(IEnumerable<string> bullets, int level = 0, char bulletChar = '\u2022', Action<PowerPointParagraph>? configure = null) {
            if (bullets == null) {
                throw new ArgumentNullException(nameof(bullets));
            }

            var results = new List<PowerPointParagraph>();
            foreach (string bullet in bullets) {
                results.Add(AddBullet(bullet ?? string.Empty, bulletChar, level, configure));
            }

            return results;
        }

        /// <summary>
        ///     Replaces all table-cell paragraphs with a bulleted list.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetBullets(IEnumerable<string> bullets, int level = 0, char bulletChar = '\u2022', Action<PowerPointParagraph>? configure = null) {
            if (bullets == null) {
                throw new ArgumentNullException(nameof(bullets));
            }

            return ReplaceTableCellParagraphs(bullets, (bullet, templateParagraph) => {
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
        ///     Adds a numbered paragraph to the table cell.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, A.TextAutoNumberSchemeValues style, int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            PowerPointParagraph paragraph = AddParagraph(text ?? string.Empty);
            paragraph.SetNumbered(style, startAt);
            if (level > 0) {
                paragraph.Level = level;
            }

            PowerPointListParagraphDefaults.Apply(paragraph);
            configure?.Invoke(paragraph);
            return paragraph;
        }

        /// <summary>
        ///     Adds a numbered paragraph to the table cell using the default numbering style.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return AddNumberedItem(text, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        /// <summary>
        ///     Adds a numbered list to the table cell.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddNumberedList(IEnumerable<string> items, A.TextAutoNumberSchemeValues style, int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
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
        ///     Adds a numbered list to the table cell using the default numbering style.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddNumberedList(IEnumerable<string> items, int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return AddNumberedList(items, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        /// <summary>
        ///     Replaces all table-cell paragraphs with a numbered list.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetNumberedList(IEnumerable<string> items, A.TextAutoNumberSchemeValues style, int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            bool first = true;
            return ReplaceTableCellParagraphs(items, (item, templateParagraph) => {
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
        ///     Replaces all table-cell paragraphs with a numbered list using the default numbering style.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetNumberedList(IEnumerable<string> items, int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return SetNumberedList(items, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        private IReadOnlyList<PowerPointParagraph> ReplaceTableCellParagraphs<T>(IEnumerable<T> items, Func<T, A.Paragraph?, PowerPointParagraph> addParagraph) {
            A.TextBody textBody = EnsureTextBody();
            string[] discardedSoundIds = PowerPointEmbeddedSound
                .GetRelationshipIds(textBody);
            A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            textBody.RemoveAllChildren<A.Paragraph>();

            var results = new List<PowerPointParagraph>();
            foreach (T item in items) {
                results.Add(addParagraph(item, templateParagraph));
            }

            if (results.Count == 0) {
                textBody.Append(CreateEmptyParagraph(templateParagraph));
            }

            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                discardedSoundIds);

            return results;
        }
    }
}
