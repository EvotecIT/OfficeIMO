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

        private IEnumerable<A.Run> Runs => Shape.TextBody!.Elements<A.Paragraph>().SelectMany(p => p.Elements<A.Run>());

        /// <summary>
        ///     Text contained in the textbox.
        /// </summary>
        public string Text {
            get {
                TextBody? textBody = Shape.TextBody;
                if (textBody == null) {
                    return string.Empty;
                }

                List<string> paragraphs = new();
                foreach (A.Paragraph paragraph in textBody.Elements<A.Paragraph>()) {
                    paragraphs.Add(paragraph.InnerText ?? string.Empty);
                }

                if (paragraphs.Count == 0) {
                    return string.Empty;
                }

                return string.Join(Environment.NewLine, paragraphs);
            }
            set {
                string textValue = value ?? string.Empty;

                TextBody? existingTextBody = Shape.TextBody;
                if (existingTextBody == null) {
                    existingTextBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
                    Shape.TextBody = existingTextBody;
                }

                TextBody textBody = existingTextBody;
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

                MarkModified();
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

                MarkModified();
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

                MarkModified();
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

                MarkModified();
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

                MarkModified();
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
                    runProps.RemoveAllChildren<A.SolidFill>();
                    if (value != null) {
                        runProps.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                    }
                }

                MarkModified();
            }
        }

        /// <summary>
        ///     Adds a new bulleted paragraph to the textbox.
        /// </summary>
        public void AddBullet(string text) {
            A.Run run = new(new A.Text(text));
            A.Run? template = Runs.FirstOrDefault();
            if (template?.RunProperties != null) {
                run.RunProperties = (A.RunProperties)template.RunProperties.CloneNode(true);
            }

            A.Paragraph paragraph = new(
                new A.ParagraphProperties(new A.CharacterBullet() { Char = "•" }),
                run
            );
            Shape.TextBody!.AppendChild(paragraph);
            MarkModified();
        }
    }
}