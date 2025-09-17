using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
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
                var section0FirstHeader = RequireSectionHeader(section0, HeaderFooterValues.First, "Section 0 first header");
                section0FirstHeader.AddParagraph().SetText("Test Section 0 - First Header");

                var section0DefaultHeader = RequireSectionHeader(section0, HeaderFooterValues.Default, "Section 0 default header");
                section0DefaultHeader.AddParagraph().SetText("Test Section 0 - Header");

                var section0EvenHeader = RequireSectionHeader(section0, HeaderFooterValues.Even, "Section 0 even header");
                section0EvenHeader.AddParagraph().SetText("Test Section 0 - Even");

                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.AddHeadersAndFooters();

                var section1DefaultHeader = RequireSectionHeader(section1, HeaderFooterValues.Default, "Section 1 default header");
                section1DefaultHeader.AddParagraph().SetText("Test Section 1 - Header");

                section1.DifferentFirstPage = true;
                var section1FirstHeader = RequireSectionHeader(section1, HeaderFooterValues.First, "Section 1 first header");
                section1FirstHeader.AddParagraph().SetText("Test Section 1 - First Header");


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section2 = document.AddSection();
                section2.AddParagraph("Test Section2");
                section2.PageOrientation = PageOrientationValues.Landscape;
                section2.AddHeadersAndFooters();
                var section2DefaultHeader = RequireSectionHeader(section2, HeaderFooterValues.Default, "Section 2 default header");
                section2DefaultHeader.AddParagraph().SetText("Test Section 2 - Header");

                document.AddParagraph("Test Section2 - Paragraph 1");


                var section3 = document.AddSection();
                section3.AddParagraph("Test Section3");
                section3.AddHeadersAndFooters();
                var section3DefaultHeader = RequireSectionHeader(section3, HeaderFooterValues.Default, "Section 3 default header");
                section3DefaultHeader.AddParagraph().SetText("Test Section 3 - Header");


                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

                var section0DefaultHeaderCreated = RequireSectionHeader(document.Sections[0], HeaderFooterValues.Default, "Section 0 default header");
                var section1DefaultHeaderCreated = RequireSectionHeader(document.Sections[1], HeaderFooterValues.Default, "Section 1 default header");
                var section2DefaultHeaderCreated = RequireSectionHeader(document.Sections[2], HeaderFooterValues.Default, "Section 2 default header");
                var section3DefaultHeaderCreated = RequireSectionHeader(document.Sections[3], HeaderFooterValues.Default, "Section 3 default header");

                Console.WriteLine("Section 0 - Text 0: " + section0DefaultHeaderCreated.Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + section1DefaultHeaderCreated.Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + section2DefaultHeaderCreated.Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + section3DefaultHeaderCreated.Paragraphs[0].Text);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("-----");
                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);
                var section0DefaultHeaderLoaded = RequireSectionHeader(document.Sections[0], HeaderFooterValues.Default, "Section 0 default header");
                var section1DefaultHeaderLoaded = RequireSectionHeader(document.Sections[1], HeaderFooterValues.Default, "Section 1 default header");
                var section2DefaultHeaderLoaded = RequireSectionHeader(document.Sections[2], HeaderFooterValues.Default, "Section 2 default header");
                var section3DefaultHeaderLoaded = RequireSectionHeader(document.Sections[3], HeaderFooterValues.Default, "Section 3 default header");

                Console.WriteLine("Section 0 - Text 0: " + section0DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + section1DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + section2DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + section3DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("-----");
                var section1DefaultHeaderAfterLoad = RequireSectionHeader(document.Sections[1], HeaderFooterValues.Default, "Section 1 default header");
                section1DefaultHeaderAfterLoad.AddParagraph().SetText("Test Section 1 - Header-Par1");
                Console.WriteLine("Section 1 - Text 1: " + section1DefaultHeaderAfterLoad.Paragraphs[1].Text);
                document.Save(openWord);
            }
        }



    }
}
