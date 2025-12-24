using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a textbox shape.
    /// </summary>
    public class PowerPointTextBox : PowerPointShape {
        internal PowerPointTextBox(Shape shape) : base(shape) {
        }

        private Shape Shape => (Shape)Element;

        private IEnumerable<A.Run> Runs => EnsureTextBody().Elements<A.Paragraph>().SelectMany(p => p.Elements<A.Run>());

        /// <summary>
        ///     Paragraphs contained in the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> Paragraphs =>
            EnsureTextBody().Elements<A.Paragraph>().Select(p => new PowerPointParagraph(p)).ToList();

        /// <summary>
        ///     Adds a paragraph to the textbox.
        /// </summary>
        public PowerPointParagraph AddParagraph(string text = "", Action<PowerPointParagraph>? configure = null,
            Action<PowerPointTextRun>? run = null) {
            TextBody textBody = EnsureTextBody();

            A.Paragraph paragraph = new();
            A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            if (templateParagraph?.ParagraphProperties != null) {
                paragraph.ParagraphProperties = (A.ParagraphProperties)templateParagraph.ParagraphProperties.CloneNode(true);
            }

            A.Run newRun = new(new A.Text(text ?? string.Empty));
            A.RunProperties? templateRunProps = templateParagraph?
                .Elements<A.Run>()
                .Select(r => r.RunProperties)
                .FirstOrDefault(rp => rp != null);
            if (templateRunProps != null) {
                newRun.RunProperties = (A.RunProperties)templateRunProps.CloneNode(true);
            }
            paragraph.Append(newRun);

            A.EndParagraphRunProperties? templateEnd = templateParagraph?.GetFirstChild<A.EndParagraphRunProperties>();
            if (templateEnd != null) {
                paragraph.Append((A.EndParagraphRunProperties)templateEnd.CloneNode(true));
            }

            textBody.Append(paragraph);

            var wrapper = new PowerPointParagraph(paragraph);
            configure?.Invoke(wrapper);
            if (run != null) {
                var runWrapper = wrapper.Runs.FirstOrDefault() ?? wrapper.AddRun(text ?? string.Empty);
                run.Invoke(runWrapper);
            }

            return wrapper;
        }

        /// <summary>
        ///     Adds multiple paragraphs to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddParagraphs(IEnumerable<string> paragraphs,
            Action<PowerPointParagraph>? configure = null) {
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
        ///     Replaces all paragraphs with the provided content.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetParagraphs(IEnumerable<string> paragraphs,
            Action<PowerPointParagraph>? configure = null) {
            if (paragraphs == null) {
                throw new ArgumentNullException(nameof(paragraphs));
            }

            Clear();
            return AddParagraphs(paragraphs, configure);
        }

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

            Clear();
            return AddBullets(bullets, level, bulletChar, configure);
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

            Clear();
            return AddNumberedList(items, style, startAt, level, configure);
        }

        /// <summary>
        ///     Replaces all paragraphs with a numbered list using the default numbering style.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetNumberedList(IEnumerable<string> items,
            int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return SetNumberedList(items, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        /// <summary>
        ///     Applies a text style to the textbox.
        /// </summary>
        public PowerPointTextBox ApplyTextStyle(PowerPointTextStyle style, bool applyToRuns = false) {
            var paragraphs = GetParagraphsForStyling();
            foreach (PowerPointParagraph paragraph in paragraphs) {
                if (applyToRuns) {
                    IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
                    if (runs.Count == 0) {
                        runs = new List<PowerPointTextRun> { paragraph.AddRun(string.Empty) };
                    }
                    foreach (PowerPointTextRun run in runs) {
                        style.Apply(run);
                    }
                } else {
                    style.Apply(paragraph);
                }
            }

            return this;
        }

        /// <summary>
        ///     Applies a paragraph style to the textbox.
        /// </summary>
        public PowerPointTextBox ApplyParagraphStyle(PowerPointParagraphStyle style) {
            var paragraphs = GetParagraphsForStyling();
            foreach (PowerPointParagraph paragraph in paragraphs) {
                style.Apply(paragraph);
            }

            return this;
        }

        /// <summary>
        ///     Applies automatic spacing defaults to the textbox paragraphs.
        /// </summary>
        public PowerPointTextBox ApplyAutoSpacing(double lineSpacingMultiplier = 1.2,
            double? spaceBeforePoints = null, double? spaceAfterPoints = null) {
            PowerPointParagraphStyle style = new(lineSpacingMultiplier: lineSpacingMultiplier,
                spaceBeforePoints: spaceBeforePoints,
                spaceAfterPoints: spaceAfterPoints);
            return ApplyParagraphStyle(style);
        }

        /// <summary>
        ///     Removes all paragraphs from the textbox.
        /// </summary>
        public void Clear() {
            EnsureTextBody().RemoveAllChildren<A.Paragraph>();
        }

        /// <summary>
        ///     Text contained in the textbox.
        /// </summary>
        public string Text {
            get {
                var paragraphs = EnsureTextBody().Elements<A.Paragraph>()
                    .Select(p => p.InnerText ?? string.Empty)
                    .ToList();
                return paragraphs.Count == 0 ? string.Empty : string.Join(Environment.NewLine, paragraphs);
            }
            set {
                string textValue = value ?? string.Empty;
                TextBody textBody = EnsureTextBody();

                A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
                A.ParagraphProperties? templateParagraphProperties = templateParagraph?.GetFirstChild<A.ParagraphProperties>();
                A.EndParagraphRunProperties? templateEndParagraphRunProperties = templateParagraph?.GetFirstChild<A.EndParagraphRunProperties>();
                A.RunProperties? templateRunProperties = templateParagraph?
                    .Elements<A.Run>()
                    .Select(r => r.RunProperties)
                    .FirstOrDefault(rp => rp != null);

                textBody.RemoveAllChildren<A.Paragraph>();

                string[] lines = textValue.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                foreach (string line in lines) {
                    A.Paragraph paragraph = new();
                    if (templateParagraphProperties != null) {
                        paragraph.Append((A.ParagraphProperties)templateParagraphProperties.CloneNode(true));
                    }

                    A.Run run = new();
                    if (templateRunProperties != null) {
                        run.RunProperties = (A.RunProperties)templateRunProperties.CloneNode(true);
                    }

                    run.Append(new A.Text(line));
                    paragraph.Append(run);

                    if (templateEndParagraphRunProperties != null) {
                        paragraph.Append((A.EndParagraphRunProperties)templateEndParagraphRunProperties.CloneNode(true));
                    }

                    textBody.Append(paragraph);
                }
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether the text is bold.
        /// </summary>
        public bool Bold {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.Bold?.Value == true;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.Bold = value ? true : null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether the text is italic.
        /// </summary>
        public bool Italic {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.Italic?.Value == true;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.Italic = value ? true : null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets the font size in points.
        /// </summary>
        public int? FontSize {
            get {
                A.Run? run = Runs.FirstOrDefault();
                int? size = run?.RunProperties?.FontSize?.Value;
                return size != null ? size / 100 : null;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.FontSize = value != null ? value * 100 : null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets the font name.
        /// </summary>
        public string? FontName {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.RemoveAllChildren<A.LatinFont>();
                    if (value != null) {
                        runProps.Append(new A.LatinFont { Typeface = value });
                    }
                }
            }
        }

        /// <summary>
        ///     Gets or sets the text color in hexadecimal format (e.g. "FF0000").
        /// </summary>
        public string? Color {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    // preserve existing fonts so we can reapply them after the fill (schema: fill must precede fonts)
                    var latin = runProps.GetFirstChild<A.LatinFont>();
                    var ea = runProps.GetFirstChild<A.EastAsianFont>();
                    var cs = runProps.GetFirstChild<A.ComplexScriptFont>();

                    runProps.RemoveAllChildren<A.SolidFill>();
                    runProps.RemoveAllChildren<A.LatinFont>();
                    runProps.RemoveAllChildren<A.EastAsianFont>();
                    runProps.RemoveAllChildren<A.ComplexScriptFont>();

                    if (value != null) {
                        runProps.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                    }

                    if (latin != null) runProps.Append((A.LatinFont)latin.CloneNode(true));
                    if (ea != null) runProps.Append((A.EastAsianFont)ea.CloneNode(true));
                    if (cs != null) runProps.Append((A.ComplexScriptFont)cs.CloneNode(true));
                }
            }
        }

        /// <summary>
        ///     Adds a new bulleted paragraph to the textbox.
        /// </summary>
        public PowerPointParagraph AddBullet(string text) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetBullet();
            return paragraph;
        }

        /// <summary>
        ///     Adds a numbered item to the textbox.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, A.TextAutoNumberSchemeValues style, int startAt = 1) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetNumbered(style, startAt);
            return paragraph;
        }

        /// <summary>
        ///     Adds a numbered item to the textbox using the default numbering style.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, int startAt = 1) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetNumbered(startAt);
            return paragraph;
        }

        private TextBody EnsureTextBody() {
            TextBody? existingTextBody = Shape.TextBody;
            if (existingTextBody == null) {
                existingTextBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
                Shape.TextBody = existingTextBody;
            }
            return existingTextBody;
        }

        private List<PowerPointParagraph> GetParagraphsForStyling() {
            var paragraphs = Paragraphs.ToList();
            if (paragraphs.Count == 0) {
                paragraphs.Add(AddParagraph());
            }
            return paragraphs;
        }
    }
}
