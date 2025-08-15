using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_NumberingDefinition(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "NumberingDefinition.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var numbering = document.CreateNumberingDefinition();
                numbering.AddLevel(new WordListLevel(WordListLevelKind.Decimal));
                var retrieved = document.GetNumberingDefinition(numbering.AbstractNumberId);
                Console.WriteLine("Numbering levels: " + retrieved.Levels.Count);
                document.Save(openWord);
            }
        }
    }
}

