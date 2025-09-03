using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {

        internal static void Example_Word_Fluent_Sections_OrientationBackground(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with section orientation and background color via fluent API");
            string filePath = Path.Combine(folderPath, "Fluent_Sections_OrientationBackground.docx");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Section(s => s
                        .New()
                            .Orientation(PageOrientationValues.Landscape)
                            .Background("FFD700")
                            .Paragraph(p => p.Text("Landscape section with background")))
                    .End()
                    .Save(false);
            }

            Helpers.Open(filePath, openWord);
        }
    }
}
