using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fields {
        internal static void Example_DocumentWithFields(string folderPath, bool openWord) {
            Console.WriteLine("[*] Opening Document with fields");
            var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates", "partitionedFieldInstructions.docx");
            

            using (WordDocument document = WordDocument.Load(filePath)) {
                foreach(var field in document.Fields) {
                    Console.WriteLine("...Type: " + field.FieldType);
                    Console.WriteLine("...Format switch: " + field.FieldFormat);
                    Console.WriteLine("...Instruction: " + String.Join(" ", field.FieldInstructions));
                    Console.WriteLine("...Switches: " + String.Join(" ", field.FieldSwitches));
                }

                //Replace ask field with new question
                var askField = document.Fields.Last();
                askField.Remove();

                var bookmark = document.Bookmarks.ToArray()[0];
                document.AddField(WordFieldType.Ask, parameters: new List<String> { bookmark.Name.ToString(), "\"How was your day?\"", "\\d \"Thanks for asking\"" });

                var fileTarget = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents", "DocumentWithFields.docx");
                document.Save(fileTarget); 
            }
        }
    }
}
