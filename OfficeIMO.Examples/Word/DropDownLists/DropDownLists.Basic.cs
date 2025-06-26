using System;
using System.IO;
using System.Collections.Generic;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class DropDownLists {
        internal static void Example_BasicDropDownList(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a dropdown list control");
            string filePath = Path.Combine(folderPath, "DocumentWithDropDownList.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var items = new List<string> { "One", "Two", "Three" };
                document.AddParagraph("Choose: ").AddDropDownList(items, "ListAlias", "ListTag");
                document.Save(openWord);
            }
        }
    }
}
