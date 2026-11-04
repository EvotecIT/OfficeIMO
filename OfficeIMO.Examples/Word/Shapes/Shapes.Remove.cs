using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_RemoveShape(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating removing a shape");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentRemoveShape.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var p = document.AddParagraph("Shape will be removed");
                var shape = p.AddShape(OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeColor.Parse("#CCCCCC"));
                Console.WriteLine($"Shape size: {shape.Width}x{shape.Height}");

                shape.Remove();

                p.AddShape(OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeColor.Parse("#FF9900"));
                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
