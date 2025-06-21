using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Revisions {
        private static string FormatErrors(IEnumerable<ValidationErrorInfo> errors) {
            return string.Join(Environment.NewLine + Environment.NewLine,
                errors.Select(error =>
                    $"Description: {error.Description}\n" +
                    $"Id: {error.Id}\n" +
                    $"ErrorType: {error.ErrorType}\n" +
                    $"Part: {error.Part?.Uri}\n" +
                    $"Path: {error.Path?.XPath}"));
        }

        internal static void Example_TrackedChanges(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating tracked changes");
            string filePath = Path.Combine(folderPath, "TrackedChangesExample.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Original text:");
                paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Inserted text", "Codex");
                paragraph.AddDeletedText("Deleted text", "Codex");
                document.Save(false);

                var valid = document.ValidateDocument();
                if (valid.Count > 0) {
                    Console.WriteLine("Document has validation errors:");
                    Console.WriteLine(FormatErrors(valid));
                }
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AcceptRevisions();
                document.Save(openWord);
            }
        }
    }
}
