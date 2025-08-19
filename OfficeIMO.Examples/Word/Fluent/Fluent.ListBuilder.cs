using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentListBuilder(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with fluent lists");
            string filePath = Path.Combine(folderPath, "FluentListBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .List(l => l.Numbered().StartAt(3)
                                 .Item("First")
                                 .Item("Second")
                                 .Indent().Item("Second.Child"))
                    .List(l => l.Bulleted()
                                 .Item("Alpha")
                                 .Item("Beta").Indent().Item("Beta.Child").Outdent()
                                 .Item("Gamma"))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
