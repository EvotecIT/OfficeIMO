using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {

        internal static void Example_SectionsWithHeadersDefault(string folderPath, bool openWord) {

            Console.WriteLine("[*] Creating standard document with sections and headers / footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections and headers footers testing.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Settings.Language = "pl-PL";

                document.Sections[0].PageOrientation = PageOrientationValues.Portrait;
                document.AddParagraph("Test Section0");
                document.AddHeadersAndFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                document.Sections[0].Header.Default.AddParagraph().SetText("Test Section 0 - Header");
                document.Sections[0].Header.First.AddParagraph().SetText("Test Section 0 - First Header");
                document.Sections[0].Header.Even.AddParagraph().SetText("Test Section 0 - Even");

                document.Sections[0].Paragraphs[0].AddComment("Przemysław Kłys", "PK", "This should be a comment");

                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();

                document.AddSection();
                document.Sections[1].PageOrientation = PageOrientationValues.Landscape;

                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();

                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                //Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Header.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Header.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Header.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Header.Default.Paragraphs[0].Text);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Sections[1].AddHeadersAndFooters();
                document.Sections[1].Header.Default.AddParagraph().SetText("Test Section 1 - Header");
                document.Sections[1].Footer.Default.AddParagraph().SetText("Test Section 1 - Header");

                document.Sections[1].DifferentFirstPage = true;
                document.Sections[1].Header.First.AddParagraph().SetText("Test Section 1 - First Header");
                document.Sections[1].Footer.First.AddParagraph().SetText("Test Section 1 - First Footer");

                document.Sections[1].DifferentOddAndEvenPages = true;

                document.Sections[1].Header.Even.AddParagraph().SetText("Test Section 1 - Even Header");
                document.Sections[1].Footer.Even.AddParagraph().SetText("Test Section 1 - Even Footer");

                document.Settings.ProtectionPassword = "ThisIsTest";
                document.Settings.ProtectionType = DocumentProtectionValues.ReadOnly;
                document.Settings.RemoveProtection();
                document.Save(openWord);
            }
        }


    }
}
