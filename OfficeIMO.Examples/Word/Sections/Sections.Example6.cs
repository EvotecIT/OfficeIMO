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

                var section0 = document.Sections[0];
                section0.PageOrientation = PageOrientationValues.Portrait;
                section0.AddParagraph("Test Section0");

                // Default header is created on demand; keep it simple
                section0.GetOrCreateHeader(HeaderFooterValues.Default)
                       .AddParagraph().SetText("Test Section 0 - Header");

                // First/even require toggles; GetOrCreate does the rest
                section0.DifferentFirstPage = true;
                section0.GetOrCreateHeader(HeaderFooterValues.First)
                       .AddParagraph().SetText("Test Section 0 - First Header");

                section0.DifferentOddAndEvenPages = true;
                section0.GetOrCreateHeader(HeaderFooterValues.Even)
                       .AddParagraph().SetText("Test Section 0 - Even");

                document.Sections[0].Paragraphs[0].AddComment("Przemysław Kłys", "PK", "This should be a comment");

                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();

                var section1Created = document.AddSection();
                section1Created.PageOrientation = PageOrientationValues.Landscape;

                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();

                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                //Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

                //Console.WriteLine("Section 0 - Text 0: " + RequireSectionHeader(document.Sections[0], HeaderFooterValues.Default, "Section 0 default header").Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + RequireSectionHeader(document.Sections[1], HeaderFooterValues.Default, "Section 1 default header").Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + RequireSectionHeader(document.Sections[2], HeaderFooterValues.Default, "Section 2 default header").Paragraphs[0].Text);
                //Console.WriteLine("Section 3 - Text 0: " + RequireSectionHeader(document.Sections[3], HeaderFooterValues.Default, "Section 3 default header").Paragraphs[0].Text);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var section1 = document.Sections[1];
                section1.GetOrCreateHeader(HeaderFooterValues.Default)
                        .AddParagraph().SetText("Test Section 1 - Header");

                section1.GetOrCreateFooter(HeaderFooterValues.Default)
                        .AddParagraph().SetText("Test Section 1 - Header");

                section1.DifferentFirstPage = true;
                section1.GetOrCreateHeader(HeaderFooterValues.First)
                        .AddParagraph().SetText("Test Section 1 - First Header");
                section1.GetOrCreateFooter(HeaderFooterValues.First)
                        .AddParagraph().SetText("Test Section 1 - First Footer");

                section1.DifferentOddAndEvenPages = true;
                section1.GetOrCreateHeader(HeaderFooterValues.Even)
                        .AddParagraph().SetText("Test Section 1 - Even Header");
                section1.GetOrCreateFooter(HeaderFooterValues.Even)
                        .AddParagraph().SetText("Test Section 1 - Even Footer");

                document.Settings.ProtectionPassword = "ThisIsTest";
                document.Settings.ProtectionType = DocumentProtectionValues.ReadOnly;
                document.Settings.RemoveProtection();
                document.Save(openWord);
            }
        }


        // No helpers needed in examples; library provides GetOrCreateHeader/Footer.
    }
}
