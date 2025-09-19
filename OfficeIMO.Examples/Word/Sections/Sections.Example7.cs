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
                var section0 = document.Sections[0];
                section0.PageOrientation = PageOrientationValues.Landscape;
                section0.AddParagraph("Test Section0");
                // Default header
                section0.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph().SetText("Test Section 0 - Header");
                // First/Evens
                section0.DifferentFirstPage = true;
                section0.GetOrCreateHeader(HeaderFooterValues.First).AddParagraph().SetText("Test Section 0 - First Header");
                section0.DifferentOddAndEvenPages = true;
                section0.GetOrCreateHeader(HeaderFooterValues.Even).AddParagraph().SetText("Test Section 0 - Even");

                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph().SetText("Test Section 1 - Header");
                section1.DifferentFirstPage = true;
                section1.GetOrCreateHeader(HeaderFooterValues.First).AddParagraph().SetText("Test Section 1 - First Header");


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                document.AddPageBreak();


                var section2 = document.AddSection();
                section2.AddParagraph("Test Section2");
                section2.PageOrientation = PageOrientationValues.Landscape;
                section2.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph().SetText("Test Section 2 - Header");

                document.AddParagraph("Test Section2 - Paragraph 1");


                var section3 = document.AddSection();
                section3.AddParagraph("Test Section3");
                section3.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph().SetText("Test Section 3 - Header");


                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

                var section0DefaultHeaderCreated = document.Sections[0].GetOrCreateHeader(HeaderFooterValues.Default);
                var section1DefaultHeaderCreated = document.Sections[1].GetOrCreateHeader(HeaderFooterValues.Default);
                var section2DefaultHeaderCreated = document.Sections[2].GetOrCreateHeader(HeaderFooterValues.Default);
                var section3DefaultHeaderCreated = document.Sections[3].GetOrCreateHeader(HeaderFooterValues.Default);

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
                var section0DefaultHeaderLoaded = document.Sections[0].GetOrCreateHeader(HeaderFooterValues.Default);
                var section1DefaultHeaderLoaded = document.Sections[1].GetOrCreateHeader(HeaderFooterValues.Default);
                var section2DefaultHeaderLoaded = document.Sections[2].GetOrCreateHeader(HeaderFooterValues.Default);
                var section3DefaultHeaderLoaded = document.Sections[3].GetOrCreateHeader(HeaderFooterValues.Default);

                Console.WriteLine("Section 0 - Text 0: " + section0DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + section1DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + section2DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + section3DefaultHeaderLoaded.Paragraphs[0].Text);
                Console.WriteLine("-----");
                var section1DefaultHeaderAfterLoad = document.Sections[1].GetOrCreateHeader(HeaderFooterValues.Default);
                section1DefaultHeaderAfterLoad.AddParagraph().SetText("Test Section 1 - Header-Par1");
                Console.WriteLine("Section 1 - Text 1: " + section1DefaultHeaderAfterLoad.Paragraphs[1].Text);
                document.Save(openWord);
            }
        }



    }
}
