using System;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddLine(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a line shape");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithLine.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with line");
                paragraph.AddLine(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeColor.Red, OfficeIMO.OfficeLength.Points(2));
                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
