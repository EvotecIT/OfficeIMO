using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentListBuilder(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with fluent lists");
            string filePath = Path.Combine(folderPath, "FluentListBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var list1 = document.AddList(WordListStyle.Numbered);
                    int level2 = 0;
                    list1.Numbering.Levels[0].SetStartNumberingValue(3);
                    list1.AddItem("First", level2);
                    list1.AddItem("Second", level2);
                    level2 = 1;
                    list1.AddItem("Second.Child", level2);
                    level2 = 2;
                    list1.AddItem("Second.Child.Grandchild", level2);
                    level2 = 0;
                    list1.AddItem("Third", level2);
                    var list3 = document.AddList(WordListStyle.Bulleted);
                    int level4 = 0;
                    list3.AddItem("Alpha", level4);
                    list3.AddItem("Beta", level4);
                    level4++;
                    list3.AddItem("Beta.Child", level4);
                    level4 = Math.Max(0, level4 - 1);
                    list3.AddItem("Gamma", level4);
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}