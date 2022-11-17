using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fields {
        internal static void Example_DocumentWithFields(string folderPath, bool openWord) {
            Console.WriteLine("[*] Opening Document with fields");
            var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates", "DocumentWithFields.docx");
            

            using (WordDocument document = WordDocument.Load(filePath)) {

                var firstField = document.Fields[0];
                Console.WriteLine("...Type: " + firstField.FieldType);
                Console.WriteLine("...Format switch: " + firstField.FieldFormat);
                Console.WriteLine("...Switches: " + String.Join(" ",firstField.FieldSwitches));

                var tocPageRef = document.Fields[1];
                Console.WriteLine("...Type: " + tocPageRef.FieldType);
                Console.WriteLine("...Format switch: " + tocPageRef.FieldFormat);
                Console.WriteLine("...Instruction: " + String.Join(" ", tocPageRef.FieldInstructions));
                Console.WriteLine("...Switches: " + String.Join(" ", tocPageRef.FieldSwitches));

                var askField = document.Fields.Find((e)=>e.FieldType == WordFieldType.Ask);
                Console.WriteLine("...Type: " + askField.FieldType);
                Console.WriteLine("...Format switch: " + askField.FieldFormat);
                Console.WriteLine("...Switches: " + String.Join(" ", askField.FieldSwitches));

                //Replace ask field with new question
                askField.Remove();
                var bookmark = document.Bookmarks.ToArray()[0];
                document.AddField(WordFieldType.Ask, parameters: new List<String> { bookmark.Name.ToString(), "\"How was your day?\"", "\\d \"Thanks for asking\"" });

                var fileTarget = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents", "DocumentWithFields.docx");
                document.Save(fileTarget);
            }
        }
    }
}
