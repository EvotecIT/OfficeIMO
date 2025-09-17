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
                GetSectionHeaderOrThrow(section0, HeaderFooterValues.First).AddParagraph().SetText("Test Section 0 - First Header");
                GetSectionHeaderOrThrow(section0).AddParagraph().SetText("Test Section 0 - Header");
                GetSectionHeaderOrThrow(section0, HeaderFooterValues.Even).AddParagraph().SetText("Test Section 0 - Even");

                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.AddHeadersAndFooters();
                GetSectionHeaderOrThrow(section1).AddParagraph().SetText("Test Section 1 - Header");
                section1.DifferentFirstPage = true;
                GetSectionHeaderOrThrow(section1, HeaderFooterValues.First).AddParagraph().SetText("Test Section 1 - First Header");


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section2 = document.AddSection();
                section2.AddParagraph("Test Section2");
                section2.PageOrientation = PageOrientationValues.Landscape;
                section2.AddHeadersAndFooters();
                GetSectionHeaderOrThrow(section2).AddParagraph().SetText("Test Section 2 - Header");

                document.AddParagraph("Test Section2 - Paragraph 1");


                var section3 = document.AddSection();
                section3.AddParagraph("Test Section3");
                section3.AddHeadersAndFooters();
                GetSectionHeaderOrThrow(section3).AddParagraph().SetText("Test Section 3 - Header");


                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

                Console.WriteLine("Section 0 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[0]).Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[1]).Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[2]).Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[3]).Paragraphs[0].Text);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("-----");
                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);
                Console.WriteLine("Section 0 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[0]).Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[1]).Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[2]).Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + GetSectionHeaderOrThrow(document.Sections[3]).Paragraphs[0].Text);
                Console.WriteLine("-----");
                var loadedSection1Header = GetSectionHeaderOrThrow(document.Sections[1]);
                loadedSection1Header.AddParagraph().SetText("Test Section 1 - Header-Par1");
                Console.WriteLine("Section 1 - Text 1: " + loadedSection1Header.Paragraphs[1].Text);
                document.Save(openWord);
            }
        }



    }
}
