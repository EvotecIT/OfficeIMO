using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {


        internal static void Example_BasicWordWithHeaderAndFooter1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Headers and Footers 1");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers Default 1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].ColumnsSpace = 50;
                Console.WriteLine("+ Settings Zoom Preset: " + document.Settings.ZoomPreset);
                Console.WriteLine("+ Settings Zoom Percent: " + document.Settings.ZoomPercentage);

                //document.Settings.ZoomPreset = PresetZoomValues.BestFit;
                //document.Settings.ZoomPercentage = 30;

                Console.WriteLine("+ Settings Zoom Preset: " + document.Settings.ZoomPreset);
                Console.WriteLine("+ Settings Zoom Percent: " + document.Settings.ZoomPercentage);

                document.AddHeadersAndFooters();
                document.DifferentFirstPage = true;
                //document.DifferentOddAndEvenPages = false;
                //var paragraphInFooter = document.Footer.Default.InsertParagraph();
                //paragraphInFooter.Text = "This is a test on odd pages (aka default if no options are set)";

                var paragraphInHeader = document.Header.Default.AddParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";

                paragraphInHeader = document.Header.First.AddParagraph();
                paragraphInHeader.Text = "First Header / Section 0";

                //var paragraphInFooterFirst = document.Footer.First.InsertParagraph();
                //paragraphInFooterFirst.Text = "This is a test on first";

                //var count = document.Footer.First.Paragraphs.Count;

                //var paragraphInFooterOdd = document.Footer.Odd.InsertParagraph();
                //paragraphInFooterOdd.Text = "This is a test odd";


                //var paragraphHeader = document.Header.Odd.InsertParagraph();
                //paragraphHeader.Text = "Header - ODD";

                //var paragraphInFooterEven = document.Footer.Even.InsertParagraph();
                //paragraphInFooterEven.Text = "This is a test - Even";


                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                //paragraph = document.InsertPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                //paragraph = document.InsertPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                //paragraph = document.InsertPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                //paragraph = document.InsertPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 5");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                //var section2 = document.InsertSection(SectionMarkValues.NextPage);
                var section2 = document.AddSection();
                section2.AddHeadersAndFooters();
                section2.DifferentFirstPage = true;
                section2.DifferentOddAndEvenPages = true;


                // Add header to section
                //var paragraghInHeaderSection = section2.Header.First.InsertParagraph();
                //paragraghInHeaderSection.Text = "Ok, work please?";

                var paragraghInHeaderSection1 = section2.Header.Default.AddParagraph();
                paragraghInHeaderSection1.Text = "Weird shit? 1";

                paragraghInHeaderSection1 = section2.Header.First.AddParagraph();
                paragraghInHeaderSection1.Text = "Weird shit 2?";
                // paragraghInHeaderSection1.InsertText("ok?");

                paragraghInHeaderSection1 = section2.Header.Even.AddParagraph();
                paragraghInHeaderSection1.Text = "Weird shit? 3";

                paragraph = document.AddParagraph("Basic paragraph - Page 6");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                paragraph = document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 7");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;


                paragraph = document.AddParagraph("Basic paragraph - Section 3.1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                paragraph = document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Section 3.2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                paragraph = document.AddPageBreak();

                //paragraph = document.Footer.Odd.InsertParagraph();
                //paragraph.Text = "Lets see";

                // 2 section, 9 paragraphs + 7 pagebreaks = 15 paragraphs, 7 pagebreaks
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                // additional sections
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                //Console.WriteLine("+ Paragraphs section 2: " + document.Sections[0].Paragraphs.Count);
                //Console.WriteLine("+ Paragraphs section 3: " + document.Sections[0].Paragraphs.Count);
                document.Save(openWord);
            }
        }

    }
}
