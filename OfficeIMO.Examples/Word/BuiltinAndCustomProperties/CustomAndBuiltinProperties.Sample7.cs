using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CustomAndBuiltinProperties {

        /// <summary>
        /// Creates a document with a few custom properties set.
        /// </summary>
        /// <param name="folderPath">Destination folder for the document.</param>
        /// <param name="openWord">Whether to open Word after creation.</param>
        public static void Example_BasicCustomProperties(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with custom properties");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with custom properties.docx");

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

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("* Loading document...");
                Console.WriteLine("+ Custom properties: " + document.CustomProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomProperties.Keys.GetEnumerator());

                document.CustomProperties["MyName"].Value = "Przemysław Kłys";

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
