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
        /// Validates a document in memory before saving it to disk.
        /// </summary>
        public static void Example_ValidateDocument_BeforeSave() {
            Console.WriteLine("[*] Creating standard document and validate it without saving");
            using (WordDocument document = WordDocument.Create()) {
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
            }
        }
    }
}
