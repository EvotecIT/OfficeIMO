using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_CustomNumberingFluent(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom numbering using fluent API");
            string filePath = Path.Combine(folderPath, "DocumentCustomNumberingFluent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .List(l => l.Numbered().NumberFormat(NumberFormatValues.UpperRoman)
                                     .Item("First")
                                     .Item("Second"))
                    .List(l => l.Bulleted().BulletCharacter("\u2192")
                                     .Item("Step 1")
                                     .Item("Step 2"))
                    .End()
                    .Save(openWord);
            }
        }
    }
}
