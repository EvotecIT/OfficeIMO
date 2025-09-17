using System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {
        internal static void Example_SectionsWithHeaders(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with sections and headers / footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections and headers footers.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section0");
                document.AddHeadersAndFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                var section0 = document.Sections[0];
                var section0FirstHeader = Guard.NotNull(section0.Header?.First, "Section 0 should expose a first header after enabling different first page.");
                section0FirstHeader.AddParagraph().SetText("Test Section 0 - First Header");
                var section0DefaultHeader = Guard.NotNull(section0.Header?.Default, "Section 0 should expose a default header after adding headers and footers.");
                section0DefaultHeader.AddParagraph().SetText("Test Section 0 - Header");
                var section0EvenHeader = Guard.NotNull(section0.Header?.Even, "Section 0 should expose an even header after enabling different odd and even pages.");
                section0EvenHeader.AddParagraph().SetText("Test Section 0 - Even");

                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.AddHeadersAndFooters();
                var section1DefaultHeader = Guard.NotNull(section1.Header?.Default, "Section 1 should expose a default header after adding headers and footers.");
                section1DefaultHeader.AddParagraph().SetText("Test Section 1 - Header");
                section1.DifferentFirstPage = true;
                var section1FirstHeader = Guard.NotNull(section1.Header?.First, "Section 1 should expose a first header after enabling different first page.");
                section1FirstHeader.AddParagraph().SetText("Test Section 1 - First Header");


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section2 = document.AddSection();
                section2.AddParagraph("Test Section2");
                section2.PageOrientation = PageOrientationValues.Landscape;
                section2.AddHeadersAndFooters();
                var section2DefaultHeader = Guard.NotNull(section2.Header?.Default, "Section 2 should expose a default header after adding headers and footers.");
                section2DefaultHeader.AddParagraph().SetText("Test Section 2 - Header");

                document.AddParagraph("Test Section2 - Paragraph 1");


                var section3 = document.AddSection();
                section3.AddParagraph("Test Section3");
                section3.AddHeadersAndFooters();
                var section3DefaultHeader = Guard.NotNull(section3.Header?.Default, "Section 3 should expose a default header after adding headers and footers.");
                section3DefaultHeader.AddParagraph().SetText("Test Section 3 - Header");


                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

                Console.WriteLine("Section 0 - Text 0: " + Guard.NotNull(document.Sections[0].Header?.Default, "Section 0 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + Guard.NotNull(document.Sections[1].Header?.Default, "Section 1 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + Guard.NotNull(document.Sections[2].Header?.Default, "Section 2 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + Guard.NotNull(document.Sections[3].Header?.Default, "Section 3 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("-----");
                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);
                Console.WriteLine("Section 0 - Text 0: " + Guard.NotNull(document.Sections[0].Header?.Default, "Section 0 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + Guard.NotNull(document.Sections[1].Header?.Default, "Section 1 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + Guard.NotNull(document.Sections[2].Header?.Default, "Section 2 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + Guard.NotNull(document.Sections[3].Header?.Default, "Section 3 should expose a default header after adding headers and footers.").Paragraphs[0].Text);
                Console.WriteLine("-----");
                var section1Header = Guard.NotNull(document.Sections[1].Header?.Default, "Section 1 should expose a default header after adding headers and footers.");
                section1Header.AddParagraph().SetText("Test Section 1 - Header-Par1");
                Console.WriteLine("Section 1 - Text 1: " + section1Header.Paragraphs[1].Text);
                document.Save(openWord);
            }
        }



    }
}
