using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {
    internal static void Example_MultipleParagraphsViaDifferentWays(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with multiple paragraphs, with some formatting");
        string filePath = System.IO.Path.Combine(folderPath, "AdvancedParagraphs.docx");
        using (WordDocument document = WordDocument.Create()) {
            var paragraph = document.AddParagraph("This paragraph starts with some text");
            paragraph.Bold = true;
            paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

            paragraph = document.AddParagraph("1st Test Second Paragraph");

            paragraph = document.AddParagraph();
            paragraph.Text = "2nd Test Third Paragraph, ";
            paragraph.Underline = WordUnderlineStyle.None;
            var paragraph2 = paragraph.AddText("3rd continuing?");
            paragraph2.Underline = WordUnderlineStyle.Double;
            paragraph2.Bold = true;
            paragraph2.Spacing = OfficeLength.Emus((long)(200) * 635);

            document.AddParagraph().SetText("4th Fourth paragraph with text").Bold = true;

            WordParagraph paragraph1 = new WordParagraph() {
                Text = "Fifth paragraph",
                Italic = true,
                Bold = true
            };
            document.AddParagraph(paragraph1);

            paragraph = document.AddParagraph("5th Test gmarmmar, this shouldnt show up as baddly written.");
            paragraph.DoNotCheckSpellingOrGrammar = true;
            paragraph.CapsStyle = WordCapsStyle.Caps;

            paragraph = document.AddParagraph("6th Test gmarmmar, this should show up as baddly written.");
            paragraph.DoNotCheckSpellingOrGrammar = false;
            paragraph.CapsStyle = WordCapsStyle.SmallCaps;

            paragraph = document.AddParagraph("7th Highlight me on a new page?");
            paragraph.Highlight = WordHighlightColor.Yellow;
            paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(15);
            paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
            paragraph.PageBreakBefore = true;


            paragraph = document.AddParagraph("8th This text should be colored.");
            paragraph.Bold = true;
            paragraph.ColorHex = "4F48E2";
            paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(70D);


            paragraph = document.AddParagraph("This is very long line that we will use to show indentation that will work across multiple lines and more and more and even more than that. One, two, three, don't worry baby.");
            paragraph.Bold = true;
            paragraph.ColorHex = "#FF0000";
            paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(36D);
            paragraph.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(70D);


            paragraph = document.AddParagraph("9th This text should be colored and Arial.");
            paragraph.Bold = true;
            paragraph.ColorHex = "4F48E2";
            paragraph.FontFamily = "Arial";
            paragraph.VerticalCharacterAlignmentOnLine = WordVerticalCharacterAlignment.Bottom;

            paragraph = document.AddParagraph("10th This text should be colored and Tahoma.");
            paragraph.Bold = true;
            paragraph.ColorHex = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(20);
            paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(15D);

            paragraph = document.AddParagraph("12th This text should be colored and Tahoma and text direction changed");
            paragraph.Bold = true;
            paragraph.ColorHex = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(10);
            paragraph.TextDirection = WordTextDirection.TopToBottomRightToLeftRotated;

            paragraph = document.AddParagraph("Spacing Test 1");
            paragraph.Bold = true;
            paragraph.ColorHex = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(36D);

            paragraph = document.AddParagraph("Spacing Test 2");
            paragraph.Bold = true;
            paragraph.ColorHex = "4F48E2";
            paragraph.FontFamily = "Tahoma";


            paragraph = document.AddParagraph("Spacing Test 3");
            paragraph.Bold = true;
            paragraph.ColorHex = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
            paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple(6.25);

            Console.WriteLine("Found paragraphs in document: " + document.Paragraphs.Count);

            document.SaveCopy(filePath);
            if (openWord) document.OpenInApplication(filePath);
        }
    }

}
