using System;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddShapesInSectionAndHeader(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shapes in section and header");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithSectionAndHeaderShapes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var section = document.Sections[0];
                section.AddShape(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeLength.Points(25), Color.Red, Color.Black);
                section.AddShape(WordShapeType.RoundedRectangle, OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(20), Color.Yellow, Color.Purple, OfficeIMO.OfficeLength.Points(1), arcSize: 0.3);
                section.AddShapeDrawing(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(40));

                section.AddHeadersAndFooters();
                var defaultHeader = RequireDefaultSectionHeader(section, "Section 0 default header");
                defaultHeader.AddShape(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeLength.Points(20), Color.Blue, Color.Black);
                defaultHeader.AddShape(WordShapeType.RoundedRectangle, OfficeIMO.OfficeLength.Points(25), OfficeIMO.OfficeLength.Points(15), Color.Green, Color.Black, OfficeIMO.OfficeLength.Points(1), arcSize: 0.3);
                defaultHeader.AddShapeDrawing(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20));

                document.Save();
                if (openWord) document.OpenInApplication();
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
