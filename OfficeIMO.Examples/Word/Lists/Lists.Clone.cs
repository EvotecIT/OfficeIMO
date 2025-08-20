using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_CloneList(string folderPath, bool openWord) {
            Console.WriteLine("[*] Cloning list");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Cloned List.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordList list = document.AddList(WordListStyle.Numbered);
                list.RestartNumberingAfterBreak = true;
                list.Numbering.Levels[0].SetStartNumberingValue(5);
                list.AddItem("Item 1");
                list.AddItem("Item 2");

                WordList cloned = list.Clone();
                cloned.AddItem("Item 3 from clone");

                document.Save(openWord);
            }
        }
    }
}
