using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists10(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists In The Middle.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                //this will be the first paragraph in the document
                var first = document.AddParagraph("First");
                //this will be 3rd pargraph in the document
                document.AddParagraph("Last");

                // Let's add a list between first and last
                // The list "placeholder" will be added to the document, but it doesn't really matter
                var list = first.AddList(WordListStyle.Bulleted);
                // This will be added to the list, and it will be the first item in the list
                // and it will be added to the document after the first paragraph
                list.AddItem("Important", 0, first);
                // This will be added to the list, but after last paragraph
                list.AddItem("List");

                document.Save(openWord);
            }
        }
    }
}
