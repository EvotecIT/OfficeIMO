using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a textbox shape.
    /// </summary>
    public class PowerPointTextBox : PowerPointShape {
        internal PowerPointTextBox(Shape shape, Action onChanged) : base(shape, onChanged) {
        }

        private Shape Shape => (Shape)Element;

        private IEnumerable<A.Run> Runs => Shape.TextBody!.Elements<A.Paragraph>().SelectMany(p => p.Elements<A.Run>());

        /// <summary>
        ///     Text contained in the textbox.
        /// </summary>
        public string Text {
            get {
                A.Run run = Shape.TextBody!.GetFirstChild<A.Paragraph>()!.GetFirstChild<A.Run>()!;
                A.Text text = run.GetFirstChild<A.Text>()!;
                return text.Text ?? string.Empty;
            }
            set {
                A.Run run = Shape.TextBody!.GetFirstChild<A.Paragraph>()!.GetFirstChild<A.Run>()!;
                A.Text text = run.GetFirstChild<A.Text>()!;
                text.Text = value;
                NotifyChanged();
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
                NotifyChanged();
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
                NotifyChanged();
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
                NotifyChanged();
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
                NotifyChanged();
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
                NotifyChanged();
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
            NotifyChanged();
        }
    }
}