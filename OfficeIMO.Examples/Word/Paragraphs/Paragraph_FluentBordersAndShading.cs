using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {

    internal static void Example_Word_Fluent_Paragraph_BordersAndShading(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with fluent bordered and shaded paragraph");
        string filePath = Path.Combine(folderPath, "Fluent_Paragraph_BordersAndShading.docx");

        using (var document = WordDocument.Create(filePath)) {
            document.AsFluent()
                .Paragraph(p => p
                    .Text("Border and shading")
                    .Border(b => {
                        b.LeftStyle = BorderValues.Thick;
                        b.LeftColor = Color.Red;
                        b.LeftSize = 24;
                    })
                    .Shading(Color.LightGray))
                .End()
                .Save(false);
        }
        Helpers.Open(filePath, openWord);
    }
}
