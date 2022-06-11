using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class CustomAndBuiltinProperties {
        public static void Example_BasicDocumentProperties(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some properties and single paragraph");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocument.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";

                var paragraph = document.AddParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.Save(openWord);
            }
        }
    }
}
