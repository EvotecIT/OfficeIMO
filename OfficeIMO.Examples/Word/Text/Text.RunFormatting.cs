using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class TextExamples {
        internal static void Example_RunFormatting(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with run-level formatting");
            string filePath = Path.Combine(folderPath, "TextRunFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p
                        .Text("Outline ", t => t.Outline())
                        .Text("Shadow ", t => t.Shadow())
                        .Text("Emboss ", t => t.Emboss())
                        .Text("SmallCaps ", t => t.SmallCaps())
                        .Text("Combined", t => t.Outline().Shadow().Emboss().SmallCaps()))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
