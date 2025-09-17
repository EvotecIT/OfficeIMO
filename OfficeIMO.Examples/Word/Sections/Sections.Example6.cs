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

                var section0 = document.Sections[0];
                var section0DefaultHeader = RequireSectionHeader(section0, HeaderFooterValues.Default, "Section 0 default header");
                section0DefaultHeader.AddParagraph().SetText("Test Section 0 - Header");

                var section0FirstHeader = RequireSectionHeader(section0, HeaderFooterValues.First, "Section 0 first header");
                section0FirstHeader.AddParagraph().SetText("Test Section 0 - First Header");

                var section0EvenHeader = RequireSectionHeader(section0, HeaderFooterValues.Even, "Section 0 even header");
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

                //Console.WriteLine("Section 0 - Text 0: " + RequireSectionHeader(document.Sections[0], HeaderFooterValues.Default, "Section 0 default header").Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + RequireSectionHeader(document.Sections[1], HeaderFooterValues.Default, "Section 1 default header").Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + RequireSectionHeader(document.Sections[2], HeaderFooterValues.Default, "Section 2 default header").Paragraphs[0].Text);
                //Console.WriteLine("Section 3 - Text 0: " + RequireSectionHeader(document.Sections[3], HeaderFooterValues.Default, "Section 3 default header").Paragraphs[0].Text);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var section1 = document.Sections[1];
                var section1DefaultHeader = RequireSectionHeader(section1, HeaderFooterValues.Default, "Section 1 default header");
                section1DefaultHeader.AddParagraph().SetText("Test Section 1 - Header");

                var section1DefaultFooter = RequireSectionFooter(section1, HeaderFooterValues.Default, "Section 1 default footer");
                section1DefaultFooter.AddParagraph().SetText("Test Section 1 - Header");

                section1.DifferentFirstPage = true;

                var section1FirstHeader = RequireSectionHeader(section1, HeaderFooterValues.First, "Section 1 first header");
                section1FirstHeader.AddParagraph().SetText("Test Section 1 - First Header");

                var section1FirstFooter = RequireSectionFooter(section1, HeaderFooterValues.First, "Section 1 first footer");
                section1FirstFooter.AddParagraph().SetText("Test Section 1 - First Footer");

                section1.DifferentOddAndEvenPages = true;

                var section1EvenHeader = RequireSectionHeader(section1, HeaderFooterValues.Even, "Section 1 even header");
                section1EvenHeader.AddParagraph().SetText("Test Section 1 - Even Header");

                var section1EvenFooter = RequireSectionFooter(section1, HeaderFooterValues.Even, "Section 1 even footer");
                section1EvenFooter.AddParagraph().SetText("Test Section 1 - Even Footer");

                document.Settings.ProtectionPassword = "ThisIsTest";
                document.Settings.ProtectionType = DocumentProtectionValues.ReadOnly;
                document.Settings.RemoveProtection();
                document.Save(openWord);
            }
        }


        private static WordHeader RequireSectionHeader(WordSection section, HeaderFooterValues type, string description) {
            if (section == null) {
                throw new ArgumentNullException(nameof(section));
            }

            EnsureSectionHeadersAndFooters(section);

            var headers = section.Header;
            if (headers == null) {
                throw new InvalidOperationException($"{description} are not available.");
            }

            WordHeader? header;
            if (type == HeaderFooterValues.Default) {
                header = headers.Default;
            } else if (type == HeaderFooterValues.Even) {
                header = headers.Even;
            } else if (type == HeaderFooterValues.First) {
                header = headers.First;
            } else {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
            }

            if (header == null) {
                throw new InvalidOperationException($"{description} is not available.");
            }

            return header;
        }

        private static WordFooter RequireSectionFooter(WordSection section, HeaderFooterValues type, string description) {
            if (section == null) {
                throw new ArgumentNullException(nameof(section));
            }

            EnsureSectionHeadersAndFooters(section);

            var footers = section.Footer;
            if (footers == null) {
                throw new InvalidOperationException($"{description} are not available.");
            }

            WordFooter? footer;
            if (type == HeaderFooterValues.Default) {
                footer = footers.Default;
            } else if (type == HeaderFooterValues.Even) {
                footer = footers.Even;
            } else if (type == HeaderFooterValues.First) {
                footer = footers.First;
            } else {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");
            }

            if (footer == null) {
                throw new InvalidOperationException($"{description} is not available.");
            }

            return footer;
        }

        private static void EnsureSectionHeadersAndFooters(WordSection section) {
            if (section.Header == null || section.Footer == null) {
                section.AddHeadersAndFooters();
            }
        }
    }
}
