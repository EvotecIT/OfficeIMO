using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CustomAndBuiltinProperties {

        public static void Example_BasicCustomProperties(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with custom properties");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with custom properties.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
                document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
                document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

                Console.WriteLine("+ Custom properties: " + document.CustomDocumentProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomDocumentProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomDocumentProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomDocumentProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomDocumentProperties.Keys.Count());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, false)) {
                Console.WriteLine("* Loading document...");
                Console.WriteLine("+ Custom properties: " + document.CustomDocumentProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomDocumentProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomDocumentProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomDocumentProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomDocumentProperties.Keys.GetEnumerator());

                document.CustomDocumentProperties["MyName"].Value = "Przemysław Kłys";

                document.Save(openWord);
            }
        }
    }
}
