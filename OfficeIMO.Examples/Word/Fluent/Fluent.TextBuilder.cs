using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentTextBuilder(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with fluent text");
            string filePath = Path.Combine(folderPath, "FluentTextBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Hello")
                        .Text(" World", t => t.BoldOn().ItalicOn().Color("#ff0000"))
                        .Text(" Formatting", t => t
                            .Underline(UnderlineValues.Single)
                            .Highlight(HighlightColorValues.Yellow)
                            .FontSize(18)
                            .FontFamily("Arial")
                            .CapsStyle(CapsStyle.SmallCaps)
                            .Strike()))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}