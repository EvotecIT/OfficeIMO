using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists10(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists In The Middle.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var first = document.AddParagraph("First");
                document.AddParagraph("Last");

                // Let's add a list between first and last
                var list = first.AddList(WordListStyle.Bulleted);
                list.AddItem("Important",0,first);
                list.AddItem("List");

                document.Save(openWord);
            }
        }
    }
}
