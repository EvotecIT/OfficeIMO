using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {
        internal static void Example_BasicWordWithSections(string folderPath, bool openWord) {

            Console.WriteLine("[*] Creating standard document with Sections");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Sections.docx");


            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test 1");
                var section1 = document.AddSection(SectionMarkValues.NextPage);

                document.AddParagraph("Test 2");
                var section2 = document.AddSection(SectionMarkValues.Continuous);

                document.AddParagraph("Test 3");
                var section3 = document.AddSection(SectionMarkValues.NextPage);
                section3.AddParagraph("Paragraph added to section number 3");
                section3.AddParagraph("Continue adding paragraphs to section 3");

                // 4 section, 5 paragraphs, 0 pagebreaks
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                // additional sections
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 3: " + document.Sections[3].Paragraphs.Count);

                // change same paragraph using section
                document.Sections[1].Paragraphs[0].Bold = true;
                // or Paragraphs list for the whole document
                document.Paragraphs[1].ColorHex = "7178a8";

                var paragraph = section1.AddParagraph("We missed paragraph on 1 section (2nd page)");
                var newParagraph = paragraph.AddParagraphAfterSelf();
                newParagraph.Text = "Some more text, after paragraph we just added.";
                newParagraph.Bold = true;


                Console.WriteLine("+ Paragraphs (repeated): " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks (repeated): " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections   (repeated): " + document.Sections.Count);
                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0 (repeated): " + document.Sections[0].Paragraphs.Count);
                // additional sections
                Console.WriteLine("+ Paragraphs section 1 (repeated): " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2 (repeated): " + document.Sections[2].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 3 (repeated): " + document.Sections[3].Paragraphs.Count);


                document.Save(openWord);
            }
        }


    }
}
