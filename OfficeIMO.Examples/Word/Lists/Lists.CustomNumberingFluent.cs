using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class Lists
    {
        internal static void Example_CustomNumberingFluent(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with custom numbering using fluent API");
            string filePath = Path.Combine(folderPath, "DocumentCustomNumberingFluent.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var list1 = document.AddList(WordListStyle.Numbered);
                    int level2 = 0;
                    list1.Numbering.Levels[level2].NumberFormat = WordNumberFormat.UpperRoman;
                    list1.AddItem("First", level2);
                    list1.AddItem("Second", level2);
                    var list3 = document.AddList(WordListStyle.Bulleted);
                    int level4 = 0;
                    list3.Numbering.Levels[level4].LevelText = "\u2192";
                    list3.AddItem("Step 1", level4);
                    list3.AddItem("Step 2", level4);
                    document.Save();
                }

                if (openWord)
                    document.OpenInApplication();
            }
        }
    }
}