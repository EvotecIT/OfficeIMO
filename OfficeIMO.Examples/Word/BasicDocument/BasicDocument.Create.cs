using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicEmptyWord(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document (empty)");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";
                document.Save(openWord);
            }
        }

        public static void Example_BasicWord(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with paragraph");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithParagraphs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                Console.WriteLine(SixLabors.ImageSharp.Color.Blue.ToHexColor());
                Console.WriteLine(SixLabors.ImageSharp.Color.Crimson.ToHexColor());
                Console.WriteLine(SixLabors.ImageSharp.Color.Aquamarine.ToHexColor());

                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                paragraph = document.AddParagraph("Adding another paragraph with some more text");
                paragraph.Bold = true;
                paragraph = paragraph.AddText(" , but now we also decided to add more text to this paragraph using different style");
                paragraph.Underline = UnderlineValues.DashLong;
                paragraph = paragraph.AddText(" , and we still continue adding more text to existing paragraph.");
                paragraph.Color = SixLabors.ImageSharp.Color.CornflowerBlue;

                document.Save(openWord);
            }
        }
    }
}
