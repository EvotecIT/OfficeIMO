using System;
using System.Collections.Generic;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class MailMerge {
        internal static void Example_MailMergeAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge advanced");
            string filePath = System.IO.Path.Combine(folderPath, "MailMergeAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"FirstName\"" })
                    .AddText(" ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"LastName\"" })
                    .AddText(",");

                document.AddParagraph("Your balance is ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Balance\"" })
                    .AddText(".");

                var values = new Dictionary<string, string> {
                    { "FirstName", "Jane" },
                    { "LastName", "Doe" },
                    { "Balance", "$200" }
                };

                WordMailMerge.Execute(document, values, removeFields: false);
                document.Save(openWord);
            }
        }
    }
}
