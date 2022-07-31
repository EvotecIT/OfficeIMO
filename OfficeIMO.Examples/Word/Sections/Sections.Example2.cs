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

        internal static void Example_BasicSections3WithColumns(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with sections 3 and columns");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections 2.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test 1 - Should be before 1st section").SetColor(Color.LightPink);

                var section1 = document.AddSection();
                section1.AddParagraph("This is a text in 2nd section");
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.ColumnCount = 2;
                for (int i = 0; i < 10; i++) {
                    section1.AddParagraph("Test 3 - Should be in 2nd section");
                }

                section1.AddParagraph("Test5");

                var section2 = document.AddSection();

                section2.AddParagraph("Test 2 - Should be after 2nd section").SetFontFamily("Tahoma").SetFontSize(20);

                section2.PageOrientation = PageOrientationValues.Landscape;

                //// primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);

                Console.WriteLine("+ PageOrientation section 0: " + document.Sections[0].PageOrientation);
                Console.WriteLine("+ PageOrientation section 1: " + document.Sections[1].PageOrientation);
                Console.WriteLine("+ PageOrientation section 2: " + document.Sections[2].PageOrientation);

                Console.WriteLine("+ ColumnCount section 0: " + document.Sections[0].ColumnCount);
                Console.WriteLine("+ ColumnCount section 1: " + document.Sections[1].ColumnCount);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("Loaded document information:");
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);

                var section1 = document.AddSection();
                section1.AddParagraph("This is a text in 2nd section");
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.ColumnCount = 2;
                for (int i = 0; i < 10; i++) {
                    section1.AddParagraph("Test 3 - Should be in 2nd section");
                }

                for (int i = 0; i < 11; i++) {
                    Console.WriteLine(document.Sections[3].Paragraphs[i].Text);
                }

                Console.WriteLine("+ Paragraphs section 3: " + document.Sections[3].Paragraphs.Count);

                document.Save(openWord);
            }
        }



    }
}
