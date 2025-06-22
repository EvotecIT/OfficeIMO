using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fields {
        internal static void Example_FieldFormatRoman(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with roman page number field");
            string filePath = System.IO.Path.Combine(folderPath, "FieldFormatRoman.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Current page: ").AddField(WordFieldType.Page, WordFieldFormat.roman);
                document.Save(openWord);
            }
        }

        internal static void Example_FieldFormatAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with advanced field formats");
            string filePath = System.IO.Path.Combine(folderPath, "FieldFormatAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Number as words: ").AddField(WordFieldType.Page, WordFieldFormat.CardText);
                document.AddParagraph("Number ordinal: ").AddField(WordFieldType.Page, WordFieldFormat.Ordinal);
                document.AddParagraph("Number hex: ").AddField(WordFieldType.Page, WordFieldFormat.Hex);
                document.Save(openWord);
            }
        }
    }
}
