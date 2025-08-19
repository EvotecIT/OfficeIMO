using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

internal static partial class Paragraphs {

    internal static void Example_Word_Fluent_Paragraph_TextAndFormatting(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with fluent paragraph text + formatting");
        string filePath = Path.Combine(folderPath, "Fluent_Paragraph_TextAndFormatting.docx");

        using (var document = WordDocument.Create(filePath)) {
            document.AsFluent()
                .Paragraph(p => p
                    .Text("Hello")
                    .Text(" World", t => t.BoldOn().ItalicOn().Color("#ff0000"))
                    .Text("!", t => t.BoldOn()))
                .End();

            document.Save(false);
        }
        Helpers.Open(filePath, openWord);
    }
}
