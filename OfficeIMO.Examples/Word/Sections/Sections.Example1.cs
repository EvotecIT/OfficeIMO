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

        internal static void Example_BasicSections2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with sections 2");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections 1.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test 1 - Should be before 1st section").SetColor(Color.LightPink);

                var section1 = document.AddSection();

                section1.PageOrientation = PageOrientationValues.Portrait;

                section1.AddParagraph("Test 1 - Should be after 1st section").SetFontFamily("Tahoma").SetFontSize(20);

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

                Console.WriteLine("+ Paragraphs section 0 Text: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("+ Paragraphs section 2 Text: " + document.Sections[2].Paragraphs[0].Text);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("Loaded document information:");
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);

                Console.WriteLine("+ PageOrientation section 0: " + document.Sections[0].PageOrientation);
                Console.WriteLine("+ PageOrientation section 1: " + document.Sections[1].PageOrientation);
                Console.WriteLine("+ PageOrientation section 2: " + document.Sections[2].PageOrientation);

                Console.WriteLine("+ Paragraphs section 0 Text: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("+ Paragraphs section 2 Text: " + document.Sections[2].Paragraphs[0].Text);

                var section1 = document.AddSection();
                section1.AddParagraph("Test Section4");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Section5");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Section6");
                section3.PageOrientation = PageOrientationValues.Portrait;

                Console.WriteLine("Loaded document information:");
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 3: " + document.Sections[3].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 4: " + document.Sections[4].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 5: " + document.Sections[5].Paragraphs.Count);

                Console.WriteLine("+ PageOrientation section 0: " + document.Sections[0].PageOrientation);
                Console.WriteLine("+ PageOrientation section 1: " + document.Sections[1].PageOrientation);
                Console.WriteLine("+ PageOrientation section 2: " + document.Sections[2].PageOrientation);
                Console.WriteLine("+ PageOrientation section 3: " + document.Sections[3].PageOrientation);
                Console.WriteLine("+ PageOrientation section 4: " + document.Sections[4].PageOrientation);
                Console.WriteLine("+ PageOrientation section 5: " + document.Sections[5].PageOrientation);


                section1.AddParagraph("This goes to section 4");

                Console.WriteLine("+ Paragraphs section 3 Text: " + document.Sections[3].Paragraphs[0].Text);
                Console.WriteLine("+ Paragraphs section 3 Count: " + document.Sections[3].Paragraphs.Count);

                document.Save(openWord);
            }
        }


    }
}
