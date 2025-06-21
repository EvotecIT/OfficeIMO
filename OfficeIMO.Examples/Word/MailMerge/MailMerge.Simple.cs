using System;
using System.Collections.Generic;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class MailMerge {
        internal static void Example_MailMergeSimple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge simple");
            string filePath = System.IO.Path.Combine(folderPath, "MailMergeSimple.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Dear ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"FirstName\"" })
                    .AddText(",");

                document.AddParagraph("Your order number ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"OrderId\"" })
                    .AddText(" has been processed.");

                var values = new Dictionary<string, string> {
                    { "FirstName", "John" },
                    { "OrderId", "12345" }
                };

                WordMailMerge.Execute(document, values);
                document.Save(openWord);
            }
        }
    }
}
