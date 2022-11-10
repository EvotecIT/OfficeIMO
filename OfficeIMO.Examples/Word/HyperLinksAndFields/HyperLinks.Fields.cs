using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HyperLinks {
        internal static void Example_AddingFields(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Fields");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Fields.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.AddParagraph("This is my test");

                document.AddParagraph("This is page number ").AddField(WordFieldType.Page);

                document.AddParagraph("Our title is ").AddField(WordFieldType.Title, WordFieldFormat.Caps);

                var para = document.AddParagraph("Our author is ").AddField(WordFieldType.Author);
                Console.WriteLine("----");
                Console.WriteLine(document.Fields[0].FieldFormat);
                Console.WriteLine(document.Fields[0].FieldType);
                Console.WriteLine(document.Fields[0].Field);
                Console.WriteLine("----");
                Console.WriteLine(document.Fields[1].FieldFormat);
                Console.WriteLine(document.Fields[1].FieldType);
                Console.WriteLine(document.Fields[1].Field);
                Console.WriteLine("----");
                Console.WriteLine(document.Fields[2].FieldFormat);
                Console.WriteLine(document.Fields[2].FieldType);
                Console.WriteLine(document.Fields[2].Field);
                Console.WriteLine("----");
                Console.WriteLine(document.Fields.Count);
                Console.WriteLine("----");
                document.Fields[1].Remove();
                Console.WriteLine(document.Fields.Count);
                Console.WriteLine("----");
                // document.Settings.UpdateFieldsOnOpen = true;
                document.Save(openWord);
            }
        }
    }
}
