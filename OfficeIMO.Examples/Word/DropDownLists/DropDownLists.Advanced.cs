using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class DropDownLists {
        internal static void Example_AdvancedDropDownList(string folderPath, bool openWord) {
            Console.WriteLine("[*] Reading dropdown list items");
            string filePath = Path.Combine(folderPath, "DocumentWithDropDownList.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                var list = document.GetDropDownListByTag("ListTag");
                Console.WriteLine($"Item count: {list.Items.Count}");
                document.Save(openWord);
            }
        }
    }
}
