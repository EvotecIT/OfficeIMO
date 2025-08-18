using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a textbox shape.
    /// </summary>
    public class PPTextBox : PPShape {
        internal PPTextBox(Shape shape) : base(shape) {
        }

        private Shape Shape => (Shape)Element;

        /// <summary>
        /// Text contained in the textbox.
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
            }
        }

        /// <summary>
        /// Adds a new bulleted paragraph to the textbox.
        /// </summary>
        public void AddBullet(string text) {
            A.Paragraph paragraph = new(
                new A.ParagraphProperties(new A.CharacterBullet()),
                new A.Run(new A.Text(text))
            );
            Shape.TextBody!.AppendChild(paragraph);
        }
    }
}

