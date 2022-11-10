using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {

        internal static void Example_SectionsWithParagraphs(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with sections 4");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections 3.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section0").SetColor(Color.LightPink);

                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;

                section1.AddParagraph("Test Section1").SetFontFamily("Tahoma").SetFontSize(20);

                var section2 = document.AddSection();

                section2.AddParagraph("Test Section2").SetFontFamily("Tahoma").SetFontSize(20);

                section2.PageOrientation = PageOrientationValues.Landscape;


                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("-----");
                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);


                var section3 = document.AddSection();
                section3.AddParagraph("Test Section3");
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);
                section3.AddParagraph("Test Section3-Par1");
                Console.WriteLine("Section 3 - Text 1: " + document.Sections[3].Paragraphs[1].Text);
                var section4 = document.AddSection();
                section4.AddParagraph("Test Section4");
                var section5 = document.AddSection();
                section5.AddParagraph("Test Section5");
                section5.PageOrientation = PageOrientationValues.Portrait;

                document.AddParagraph("Test Section5-Par1");
                document.AddParagraph("Test Section5-Par2");
                section3.AddParagraph("Test Section3-Par2");

                Console.WriteLine("-----");
                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);
                Console.WriteLine("Section 4 - Text 0: " + document.Sections[4].Paragraphs[0].Text);
                Console.WriteLine("Section 5 - Text 0: " + document.Sections[5].Paragraphs[0].Text);
                Console.WriteLine("Section 5 - Text 1: " + document.Sections[5].Paragraphs[1].Text);
                Console.WriteLine("Section 5 - Text 2: " + document.Sections[5].Paragraphs[2].Text);
                Console.WriteLine("Section 3 - Text 1: " + document.Sections[3].Paragraphs[1].Text);
                Console.WriteLine("Section 3 - Text 2: " + document.Sections[3].Paragraphs[2].Text);
                document.Save(openWord);
            }
        }



    }
}
