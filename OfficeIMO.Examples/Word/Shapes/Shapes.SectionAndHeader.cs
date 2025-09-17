using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddShapesInSectionAndHeader(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shapes in section and header");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithSectionAndHeaderShapes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var section = document.Sections[0];
                section.AddShape(ShapeType.Rectangle, 50, 25, Color.Red, Color.Black);
                section.AddShape(ShapeType.RoundedRectangle, 40, 20, Color.Yellow, Color.Purple, 1, arcSize: 0.3);
                section.AddShapeDrawing(ShapeType.Ellipse, 40, 40);

                section.AddHeadersAndFooters();
                var defaultHeader = RequireDefaultSectionHeader(section, "Section 0 default header");
                defaultHeader.AddShape(ShapeType.Rectangle, 30, 20, Color.Blue, Color.Black);
                defaultHeader.AddShape(ShapeType.RoundedRectangle, 25, 15, Color.Green, Color.Black, 1, arcSize: 0.3);
                defaultHeader.AddShapeDrawing(ShapeType.Ellipse, 20, 20);

                document.Save(openWord);
            }
        }

        private static WordHeader RequireDefaultSectionHeader(WordSection section, string description) {
            if (section == null) {
                throw new ArgumentNullException(nameof(section));
            }

            if (section.Header == null) {
                section.AddHeadersAndFooters();
            }

            var headers = section.Header;
            if (headers == null) {
                throw new InvalidOperationException($"{description} headers are not available.");
            }

            var header = headers.Default;
            if (header == null) {
                throw new InvalidOperationException($"{description} is not available.");
            }

            return header;
        }
    }
}
