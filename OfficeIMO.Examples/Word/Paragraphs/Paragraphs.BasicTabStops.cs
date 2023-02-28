using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

internal static partial class Paragraphs {

    internal static void Example_BasicTabStops(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with paragraphs");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some tab stops.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("\tFirst Line");

            Console.WriteLine("Tabs count: " + paragraph.Tabs.Count);

            var tab1 = paragraph.AddTab(1440);

            var tab2 = paragraph.AddTab(1440);
            tab2.Alignment = TabStopValues.Left;
            tab2.Leader = TabStopLeaderCharValues.Hyphen;
            tab2.Position = 1440;

            paragraph.AddText("\tMore text");

            Console.WriteLine($"Tabs count: " + paragraph.Tabs.Count);

            var paragraph1 = document.AddParagraph("\tNext Line");

            var tab3 = paragraph1.AddTab(5000);
            tab3.Leader = TabStopLeaderCharValues.Hyphen;

            var tab4 = paragraph1.AddTab(1440 * 2);
            paragraph1.AddText("\tEven more text");

            Console.WriteLine("Tabs for Paragraph2 count: " + paragraph.Tabs.Count);
            Console.WriteLine("Tabs for Paragraph1 count: " + paragraph1.Tabs.Count);

            document.Save();
        }

        using (WordDocument document = WordDocument.Load(filePath)) {

            document.Save(openWord);
        }
    }

}
