using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fields {
        internal static void Example_FieldBuilderSimple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document using WordFieldBuilder");
            string filePath = System.IO.Path.Combine(folderPath, "FieldBuilderSimple.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var builder = new WordFieldBuilder(WordFieldType.Author)
                    .SetFormat(WordFieldFormat.Caps);
                document.AddParagraph("Author: ").AddField(builder);
                document.Save(openWord);
            }
        }

        internal static void Example_FieldBuilderNested(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document using nested WordFieldBuilder");
            string filePath = System.IO.Path.Combine(folderPath, "FieldBuilderNested.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var dateField = new WordFieldBuilder(WordFieldType.Date)
                    .SetCustomFormat("yyyy-MM-dd");
                var setField = new WordFieldBuilder(WordFieldType.Set)
                    .AddInstruction("currentDate")
                    .AddInstruction(dateField);
                document.AddParagraph().AddField(setField);

                var refField = new WordFieldBuilder(WordFieldType.Ref)
                    .AddInstruction("currentDate");
                document.AddParagraph("Saved on: ").AddField(refField);
                document.Save(openWord);
            }
        }
    }
}
