using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class CustomAndBuiltinProperties {
        /// <summary>
        /// Creates a document and validates its properties and content.
        /// </summary>
        /// <param name="folderPath">Destination folder for the document.</param>
        public static void Example_ValidateDocument(string folderPath) {
            Console.WriteLine("[*] Creating standard document and validate it");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document for validation.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;

                document.CustomProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
                document.CustomProperties.Add("MyName", new WordCustomProperty("Some text"));
                document.CustomProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

                Console.WriteLine("+ Custom properties: " + document.CustomProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomProperties.Keys.Count());

                Console.WriteLine(document.DocumentIsValid);
                Console.WriteLine(document.DocumentValidationErrors);
                document.Save();
            }
        }


    }
}
