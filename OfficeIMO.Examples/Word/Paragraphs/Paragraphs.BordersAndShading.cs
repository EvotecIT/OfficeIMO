using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {

    internal static void Example_BordersAndShading(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with bordered and shaded paragraphs");
        string filePath = System.IO.Path.Combine(folderPath, "Paragraphs with borders and shading.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var bordered = document.AddParagraph("Bordered paragraph");
            bordered.Borders.LeftStyle = BorderValues.Thick;
            bordered.Borders.LeftColor = Color.Red;
            bordered.Borders.LeftSize = 24;

            var shaded = document.AddParagraph("Shaded paragraph");
            shaded.ShadingFillColor = Color.LightGray;

            document.Save(openWord);
        }
    }
}
