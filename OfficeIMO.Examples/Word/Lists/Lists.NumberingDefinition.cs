using System;
using System.IO;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_NumberingDefinition(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "NumberingDefinition.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var numbering = document.CreateNumberingDefinition();
                numbering.AddLevel(new WordListLevel(WordListLevelKind.Decimal));
                var retrieved = Guard.NotNull(document.GetNumberingDefinition(numbering.AbstractNumberId), "Numbering definition should be retrievable after creation.");
                Console.WriteLine("Numbering levels: " + retrieved.Levels.Count);
                document.Save(openWord);
            }
        }
    }
}

