using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
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

                var section0 = document.Sections[0];
                var section0DefaultHeader = Guard.NotNull(section0.Header?.Default, "Section 0 should expose a default header after adding headers and footers.");
                section0DefaultHeader.AddParagraph().SetText("Test Section 0 - Header");
                var section0FirstHeader = Guard.NotNull(section0.Header?.First, "Section 0 should expose a first header after enabling different first page.");
                section0FirstHeader.AddParagraph().SetText("Test Section 0 - First Header");
                var section0EvenHeader = Guard.NotNull(section0.Header?.Even, "Section 0 should expose an even header after enabling different odd and even pages.");
                section0EvenHeader.AddParagraph().SetText("Test Section 0 - Even");

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

                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Header!.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Header!.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Header!.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Header!.Default.Paragraphs[0].Text);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Sections[1].AddHeadersAndFooters();
                var section1 = document.Sections[1];
                var section1DefaultHeader = Guard.NotNull(section1.Header?.Default, "Section 1 should expose a default header after adding headers and footers.");
                section1DefaultHeader.AddParagraph().SetText("Test Section 1 - Header");
                var section1DefaultFooter = Guard.NotNull(section1.Footer?.Default, "Section 1 should expose a default footer after adding headers and footers.");
                section1DefaultFooter.AddParagraph().SetText("Test Section 1 - Header");

                document.Sections[1].DifferentFirstPage = true;
                var section1FirstHeader = Guard.NotNull(section1.Header?.First, "Section 1 should expose a first header after enabling different first page.");
                section1FirstHeader.AddParagraph().SetText("Test Section 1 - First Header");
                var section1FirstFooter = Guard.NotNull(section1.Footer?.First, "Section 1 should expose a first footer after enabling different first page.");
                section1FirstFooter.AddParagraph().SetText("Test Section 1 - First Footer");

                document.Sections[1].DifferentOddAndEvenPages = true;

                var section1EvenHeader = Guard.NotNull(section1.Header?.Even, "Section 1 should expose an even header after enabling different odd and even pages.");
                section1EvenHeader.AddParagraph().SetText("Test Section 1 - Even Header");
                var section1EvenFooter = Guard.NotNull(section1.Footer?.Even, "Section 1 should expose an even footer after enabling different odd and even pages.");
                section1EvenFooter.AddParagraph().SetText("Test Section 1 - Even Footer");

                document.Settings.ProtectionPassword = "ThisIsTest";
                document.Settings.ProtectionType = DocumentProtectionValues.ReadOnly;
                document.Settings.RemoveProtection();
                document.Save(openWord);
            }
        }


    }
}
