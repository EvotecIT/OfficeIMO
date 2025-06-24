using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_ListStartNumber(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document list starting number.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var list = document.AddCustomList();
                var level = new WordListLevel(WordListLevelKind.Decimal)
                    .SetStartNumberingValue(3);
                list.Numbering.AddLevel(level);
                list.AddItem("Starts at three");

                document.Save(openWord);
            }
        }
    }
}
