using System;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_LoadShapes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading document and reading shapes");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithShapesToLoad.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Shape one").AddShape(90, 40, "#00FFFF");
                document.AddParagraph("Shape two").AddShape(60, 60, "#FF00FF");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                foreach (var paragraph in document.Paragraphs) {
                    if (paragraph.IsShape) {
                        var shape = Guard.NotNull(paragraph.Shape, "Paragraph flagged as shape should expose shape information.");
                        Console.WriteLine($"Found shape {shape.Width}x{shape.Height}");
                    }
                }
                document.Save(openWord);
            }
        }
    }
}
