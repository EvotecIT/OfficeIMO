using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_CloneList(string folderPath, bool openWord) {
            Console.WriteLine("[*] Cloning list");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Cloned List.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordList list = document.AddList(WordListStyle.Bulleted);
                list.AddItem("Item 1");
                list.AddItem("Item 2");

                WordList cloned = list.Clone();
                cloned.ListItems[0].Bold = true;

                document.Save(openWord);
            }
        }
    }
}
