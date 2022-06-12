using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWord2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with paragraph (2)");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithParagraphs2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Settings.ZoomPercentage = 50;
                var paragraph = document.AddParagraph("Basic paragraph");

                var section1 = document.AddSection();
                section1.AddParagraph("Test Middle Section - 1");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Last Section - 1");
                section1.AddParagraph("Test Middle Section - 2").AddComment("Adam Kłys", "AK", "Another test");
                var test = document.AddParagraph("Test 1 - to delete");
                test.Remove();
                section1.PageSettings.PageSize = WordPageSize.A5;
                section2.PageOrientation = PageOrientationValues.Landscape;

                document.Sections[2].AddParagraph("Test 0 - Section Last");
                document.Sections[1].AddParagraph("Test 1").AddComment("Przemysław Kłys", "PK", " This is just a test");

                Console.WriteLine("----");
                Console.WriteLine("Sections: " + document.Sections.Count);
                Console.WriteLine("----");
                Console.WriteLine(document.Sections[0].Paragraphs.Count);
                Console.WriteLine(document.Sections[1].Paragraphs.Count);
                Console.WriteLine(document.Sections[2].Paragraphs.Count);

                Console.WriteLine(document.Comments.Count);

                document.Comments[0].Text = "Lets change it";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("----");
                Console.WriteLine(document.Sections.Count);
                Console.WriteLine("----");
                Console.WriteLine(document.Sections[0].Paragraphs.Count);
                Console.WriteLine(document.Sections[0].Paragraphs.Count);
                Console.WriteLine(document.Sections[0].Paragraphs.Count);

                Console.WriteLine(document.Sections[0].HyperLinks.Count);
                Console.WriteLine(document.HyperLinks.Count);
                Console.WriteLine(document.Fields.Count);
                document.Save(openWord);
            }
        }

        public static void Example_BasicDocumentWithoutUsing(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document without using");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocumentWithSingleParagraph.docx");
            WordDocument document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";

            document.AddParagraph("This is my test");

            document.Save();
            document.Dispose();

            Helpers.Open(filePath, openWord);

            Console.WriteLine("+ IsLocked " + filePath.IsFileLocked());
        }
    }
}
