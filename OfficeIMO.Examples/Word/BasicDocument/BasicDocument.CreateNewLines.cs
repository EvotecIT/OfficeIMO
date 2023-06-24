using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithNewLines(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with different default style (PL)");
            string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithTabs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var paragraph1 = document.AddParagraph("This is a start");

                Console.WriteLine("Paragraph count (expected 1): " + document.Paragraphs.Count);

                var paragraph2 = document.AddParagraph("This is a start \t\t And more");

                Console.WriteLine("Paragraph count (expected 2): " + document.Paragraphs.Count);

                var paragraph3 = document.AddParagraph("This is a start \t\t And more");
                paragraph3.Underline = UnderlineValues.DashLong;

                Console.WriteLine("Paragraph count (expected 3): " + document.Paragraphs.Count);

                // now we will try to add new line characters, that will force the paragraph to be split into pieces
                // following will split the paragraph into 3 paragraphs - 2 paragraphs with Text, and one Break()
                var paragraph4 = document.AddParagraph("First line\r\nAnd more in new line");

                Console.WriteLine("Paragraph count (expected 6): " + document.Paragraphs.Count);

                // following will split the paragraph into 3 paragraphs - 2 paragraphs with Text, and one Break()
                var paragraph6 = document.AddParagraph("First line\nnd more in new line");

                Console.WriteLine("Paragraph count (expected 9): " + document.Paragraphs.Count);

                // following will split the paragraph into 3 paragraphs - 2 paragraphs with Text, and one Break()
                var paragraph7 = document.AddParagraph("First line" + Environment.NewLine + "And more in new line");

                Console.WriteLine("Paragraph count (expected 12): " + document.Paragraphs.Count);

                // following will split the paragraph into 7 paragraphs - 3 paragraphs with Text, and 4 paragraphs with Break()
                // additionally there's one paragraph at start
                var paragraph8 = document.AddParagraph("TestMe").AddText("\nFirst line\r\nAnd more " + Environment.NewLine + "in new line\r\n");

                Console.WriteLine("Paragraph count (expected 20): " + document.Paragraphs.Count);

                // following will split the paragraph into 7 paragraphs - 3 paragraphs with Text, and 4 paragraphs with Break()
                // additionally there's one paragraph at start
                // it's the same as above but written in a direct way. All above methods are just shortcuts for this
                var paragraph9 = document.AddParagraph("TestMe").AddBreak().AddText("First line").AddBreak().AddText("And more ").AddBreak().AddText("in new line").AddBreak();

                Console.WriteLine("Paragraph count (expected 28): " + document.Paragraphs.Count);

                document.Save(openWord);
            }
        }
    }
}
