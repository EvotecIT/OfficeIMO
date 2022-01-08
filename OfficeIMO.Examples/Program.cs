using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Helper;
using System;
using System.IO;
using System.Linq;
using Color = System.Drawing.Color;

namespace OfficeIMO.Examples {
    internal static class Program {
        private static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            } else {
               // Directory.Delete(path, true);
               // Directory.CreateDirectory(path);
            }
        }

        static void Main(string[] args) {
            //string folderPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "documents");
            string templatesPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            string folderPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            Setup(folderPath);
            string filePath;

            //Console.WriteLine("[*] Creating standard document (empty)");
            //string filePath = System.IO.Path.Combine(folderPath, "EmptyDocument.docx");
            //Example_BasicEmptyWord(filePath, false);

            //Console.WriteLine("[*] Creating standard document with paragraph");
            //filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithParagraphs.docx");
            //Example_BasicWord(filePath, true);

            //Console.WriteLine("[*] Creating standard document with some properties and single paragraph");
            //filePath = System.IO.Path.Combine(folderPath, "BasicDocument.docx");
            //Example_BasicDocumentProperties(filePath, false);

            //Console.WriteLine("[*] Creating standard document with multiple paragraphs, with some formatting");
            //filePath = System.IO.Path.Combine(folderPath, "AdvancedParagraphs.docx");
            //Example_MultipleParagraphsViaDifferentWays(filePath, false);

            //Console.WriteLine("[*] Creating standard document with some Images");
            //filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImages.docx");
            //Example_AddingImages(filePath, false);

            //Console.WriteLine("[*] Read Basic Word");
            //Example_ReadWord(true);

            //Console.WriteLine("[*] Read Basic Word with Images");
            //Example_ReadWordWithImages();

            //Console.WriteLine("[*] Creating standard document with page breaks and removing them");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with some page breaks.docx");
            //Example_PageBreaks(filePath, true);

            //Console.WriteLine("[*] Creating standard document with page breaks and removing them");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with some page breaks1.docx");
            //Example_PageBreaks1(filePath, true);

            //Console.WriteLine("[*] Creating standard document with sections");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections.docx");
            //Example_BasicSections(filePath, true);

            //Console.WriteLine("[*] Creating standard document with Sections");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with Sections.docx");
            //Example_BasicWordWithSections(filePath, true);

            //Console.WriteLine("[*] Creating standard document with Headers and Footers");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers.docx");
            //Example_BasicWordWithHeaderAndFooterWithoutSections(filePath, false);

            //Console.WriteLine("[*] Creating standard document with Page Orientation");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with PageOrientationChange.docx");
            //Example_PageOrientation(filePath, true);

            //Console.WriteLine("[*] Creating standard document with Headers and Footers including Sections");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers.docx");
            //Example_BasicWordWithHeaderAndFooter(filePath, true);

            //Console.WriteLine("[*] Loading basic document");
            //Example_Load(filePath, true);

            //Console.WriteLine("[*] Creating standard document with paragraphs");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with some paragraphs.docx");
            //Example_BasicParagraphs(filePath, true);

            //Console.WriteLine("[*] Creating standard document with custom properties");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document with custom properties.docx");
            //Example_BasicCustomProperties(filePath, true);

            //Console.WriteLine("[*] Creating standard document and validate it");
            //filePath = System.IO.Path.Combine(folderPath, "Basic Document for validation.docx");
            //Example_ValidateDocument(filePath);

            //Console.WriteLine("[*] Creating standard document and validate it without saving");
            //Example_ValidateDocument_BeforeSave();

            //Console.WriteLine("[*] Loading standard document to check properties");
            //Example_LoadDocumentWithProperties(true);

            Console.WriteLine("[*] Creating standard document with paragraphs");
            filePath = System.IO.Path.Combine(folderPath, "Document with Lists1.docx");
            Example_BasicLists(filePath, true);
            filePath = System.IO.Path.Combine(folderPath, "Document with Lists2.docx");
            Example_BasicLists2(filePath, true);
        }
        
        private static void Example_BasicEmptyWord(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";
                document.Save(openWord);
            }
        }

        private static void Example_BasicWord(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.InsertParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                paragraph = document.InsertParagraph("2nd paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AppendText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;
                paragraph = paragraph.AppendText("More text");
                paragraph.Color = System.Drawing.Color.CornflowerBlue.ToHexColor();

                document.Save(openWord);
            }
        }

        private static void Example_BasicDocumentProperties(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";

                var paragraph = document.InsertParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.Save(openWord);
            }
        }

        private static void Example_MultipleParagraphsViaDifferentWays(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create()) {
                var paragraph = document.InsertParagraph("This paragraph starts with some text");
                paragraph.Bold = true;
                paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

                paragraph = document.InsertParagraph("1st Test Second Paragraph");

                paragraph = document.InsertParagraph();
                paragraph.Text = "2nd Test Third Paragraph, ";
                paragraph.Underline = UnderlineValues.None;
                var paragraph2 = paragraph.AppendText("3rd continuing?");
                paragraph2.Underline = UnderlineValues.Double;
                paragraph2.Bold = true;
                paragraph2.Spacing = 200;

                document.InsertParagraph().InsertText("4th Fourth paragraph with text").Bold = true;

                WordParagraph paragraph1 = new WordParagraph() {
                    Text = "Fifth paragraph",
                    Italic = true,
                    Bold = true
                };
                document.InsertParagraph(paragraph1);

                paragraph = document.InsertParagraph("5th Test gmarmmar, this shouldnt show up as baddly written.");
                paragraph.DoNotCheckSpellingOrGrammar = true;
                paragraph.CapsStyle = CapsStyle.Caps;

                paragraph = document.InsertParagraph("6th Test gmarmmar, this should show up as baddly written.");
                paragraph.DoNotCheckSpellingOrGrammar = false;
                paragraph.CapsStyle = CapsStyle.SmallCaps;

                paragraph = document.InsertParagraph("7th Highlight me?");
                paragraph.Highlight = HighlightColorValues.Yellow;
                paragraph.FontSize = 15;
                paragraph.ParagraphAlignment = JustificationValues.Center;


                paragraph = document.InsertParagraph("8th This text should be colored.");
                paragraph.Bold = true;
                paragraph.Color = "4F48E2";
                paragraph.IndentationAfter = 1400;


                paragraph = document.InsertParagraph("This is very long line that we will use to show indentation that will work across multiple lines and more and more and even more than that. One, two, three, don't worry baby.");
                paragraph.Bold = true;
                paragraph.Color = "#FF0000";
                paragraph.IndentationBefore = 720;
                paragraph.IndentationFirstLine = 1400;


                paragraph = document.InsertParagraph("9th This text should be colored and Arial.");
                paragraph.Bold = true;
                paragraph.Color = "4F48E2";
                paragraph.FontFamily = "Arial";
                paragraph.VerticalCharacterAlignmentOnLine = VerticalTextAlignmentValues.Bottom;

                paragraph = document.InsertParagraph("10th This text should be colored and Tahoma.");
                paragraph.Bold = true;
                paragraph.Color = "4F48E2";
                paragraph.FontFamily = "Tahoma";
                paragraph.FontSize = 20;
                paragraph.LineSpacingBefore = 300;

                paragraph = document.InsertParagraph("12th This text should be colored and Tahoma and text direction changed");
                paragraph.Bold = true;
                paragraph.Color = "4F48E2";
                paragraph.FontFamily = "Tahoma";
                paragraph.FontSize = 10;
                paragraph.TextDirection = TextDirectionValues.TopToBottomRightToLeftRotated;

                paragraph = document.InsertParagraph("Spacing Test 1");
                paragraph.Bold = true;
                paragraph.Color = "4F48E2";
                paragraph.FontFamily = "Tahoma";
                paragraph.LineSpacingAfter = 720;

                paragraph = document.InsertParagraph("Spacing Test 2");
                paragraph.Bold = true;
                paragraph.Color = "4F48E2";
                paragraph.FontFamily = "Tahoma";


                paragraph = document.InsertParagraph("Spacing Test 3");
                paragraph.Bold = true;
                paragraph.Color = "4F48E2";
                paragraph.FontFamily = "Tahoma";
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.LineSpacing = 1500;

                Console.WriteLine("Found paragraphs in document: " + document.Paragraphs.Count);

                document.Save(filePath, openWord);
            }
        }

        private static void Example_AddingImages(string filePath, bool openWord) {
            //string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            //string imagePaths = System.IO.Path.Combine(baseDirectory, "Images");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            WordDocument document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "This is sparta";
            document.BuiltinDocumentProperties.Creator = "Przemek";

            var paragraph = document.InsertParagraph("This paragraph starts with some text");
            paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

            // lets add image to paragraph
            paragraph.InsertImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
            //paragraph.Image.WrapText = true; // WrapSideValues.Both;

            var paragraph5 = paragraph.AppendText("and more text");
            paragraph5.Bold = true;


            document.InsertParagraph("This adds another picture with 500x500");

            var filePathImage = System.IO.Path.Combine(imagePaths, "Kulek.jpg");
            WordParagraph paragraph2 = document.InsertParagraph();
            paragraph2.InsertImage(filePathImage, 500, 500);
            //paragraph2.Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
            paragraph2.Image.Rotation = 180;
            paragraph2.Image.Shape = ShapeTypeValues.ActionButtonMovie;


            document.InsertParagraph("This adds another picture with 100x100");

            WordParagraph paragraph3 = document.InsertParagraph();
            paragraph3.InsertImage(filePathImage, 100, 100);

            // we add paragraph with an image
            WordParagraph paragraph4 = document.InsertParagraph();
            paragraph4.InsertImage(filePathImage);

            // we can get the height of the image from paragraph
            Console.WriteLine("This document has image, which has height of: " + paragraph4.Image.Height + " pixels (I think) ;-)");

            // we can also overwrite height later on
            paragraph4.Image.Height = 50;
            paragraph4.Image.Width = 50;
            // this doesn't work
            paragraph4.Image.HorizontalFlip = true;

            // or we can get any image and overwrite it's size
            document.Images[0].Height = 200;
            document.Images[0].Width = 200;

            string fileToSave = System.IO.Path.Combine(imagePaths, "OutputPrzemyslawKlysAndKulkozaurr.jpg");
            document.Images[0].SaveToFile(fileToSave);

            document.Save(true);
        }

        private static void Example_ReadWord(bool openWord) {
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "BasicDocument.docx"), true);

            Console.WriteLine("This document has " + document.Paragraphs.Count + " paragraphs. Cool right?");
            Console.WriteLine("+ Document Title: " + document.BuiltinDocumentProperties.Title);
            Console.WriteLine("+ Document Author: " + document.BuiltinDocumentProperties.Creator);
            Console.WriteLine("+ FileOpen: " + document.FileOpenAccess);
            
            document.Dispose();
        }

        private static void Example_ReadWordWithImages() {
            string outputPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "BasicDocumentWithImages.docx"), true);
            Console.WriteLine("+ Document paragraphs: " + document.Paragraphs.Count);
            Console.WriteLine("+ Document images: " + document.Images.Count);

            document.Images[0].SaveToFile(System.IO.Path.Combine(outputPath, "random.jpg"));
        }

        private static void Example_PageBreaks(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";

                var paragraph = document.InsertParagraph("Test 1");

                //paragraph = new WordParagraph(document);
                //WordSection section = new WordSection(document, paragraph);


                //document._document.Body.Append(PageBreakParagraph);
                //document._document.Body.InsertBefore(PageBreakParagraph, paragraph._paragraph);

                document.InsertPageBreak();

                paragraph.Text = "Test 2";

                paragraph = document.InsertParagraph("Test 2");

                // Now lets remove paragraph with page break
                document.Paragraphs[1].Remove();

                // Now lets remove 1st paragraph
                document.Paragraphs[0].Remove();

                document.InsertPageBreak();

                document.InsertParagraph().Text = "Some text on next page";

                var paragraph1 = document.InsertParagraph("Test").AppendText("Test2");
                paragraph1.Color = System.Drawing.Color.Red.ToHexColor();
                paragraph1.AppendText("Test3");

                paragraph = document.InsertParagraph("Some paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AppendText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                paragraph = document.InsertParagraph("2nd paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AppendText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;
                paragraph = paragraph.AppendText(" More text");
                paragraph.Color = System.Drawing.Color.CornflowerBlue.ToHexColor();

                // remove last paragraph
                document.Paragraphs.Last().Remove();

                paragraph = document.InsertParagraph("2nd paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AppendText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;
                paragraph = paragraph.AppendText(" More text");
                paragraph.Color = System.Drawing.Color.CornflowerBlue.ToHexColor();

                // remove paragraph
                int countParagraphs = document.Paragraphs.Count;
                document.Paragraphs[countParagraphs - 2].Remove();

                // remove first page break
                document.PageBreaks[0].Remove();

                document.Save(openWord);
            }
        }

        private static void Example_PageBreaks1(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {

                var paragraph = document.InsertParagraph("Test 1");
                paragraph.Text = "Test 2";

                document.InsertPageBreak();

                document.InsertPageBreak();

                var paragraph1 = document.InsertParagraph("Test 1");
                paragraph1.Text = "Test 3";


                document.Save(openWord);
            }
        }

        private static void Example_BasicWordWithHeaderAndFooter1(string filePath, bool openWord) {
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

                var paragraphInHeader = document.Header.Default.InsertParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";

                paragraphInHeader = document.Header.First.InsertParagraph();
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


                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                //paragraph = document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                //paragraph = document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                //paragraph = document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                //paragraph = document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 5");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                //var section2 = document.InsertSection(SectionMarkValues.NextPage);
                var section2 = document.InsertSection();
                section2.AddHeadersAndFooters();
                section2.DifferentFirstPage = true;
                

                // Add header to section
                //var paragraghInHeaderSection = section2.Header.First.InsertParagraph();
                //paragraghInHeaderSection.Text = "Ok, work please?";

                var paragraghInHeaderSection1 = section2.Header.Default.InsertParagraph();
                paragraghInHeaderSection1.Text = "Weird shit? 1";

                paragraghInHeaderSection1 = section2.Header.First.InsertParagraph();
                paragraghInHeaderSection1.Text = "Weird shit 2?";
               // paragraghInHeaderSection1.InsertText("ok?");

                paragraghInHeaderSection1 = section2.Header.Even.InsertParagraph();
                paragraghInHeaderSection1.Text = "Weird shit? 3";

                paragraph = document.InsertParagraph("Basic paragraph - Page 6");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                paragraph = document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 7");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();


                paragraph = document.InsertParagraph("Basic paragraph - Section 3.1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                paragraph = document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Section 3.2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                paragraph = document.InsertPageBreak();

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

        private static void Example_PageOrientation(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                Console.WriteLine("+ Page Orientation (starting): " + document.PageOrientation);

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;

                Console.WriteLine("+ Page Orientation (middle): " + document.PageOrientation);
                
                document.PageOrientation = PageOrientationValues.Portrait;

                Console.WriteLine("+ Page Orientation (ending): " + document.PageOrientation);

                document.InsertParagraph("Test");
                
                document.Save(openWord);
            }
        }

        private static void Example_BasicWordWithHeaderAndFooterWithoutSections(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.BuiltinDocumentProperties.Title = "This is a test for Title";
                document.BuiltinDocumentProperties.Category = "This is a test for Category";
                
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;



                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var paragraphInHeaderO = document.Header.Default.InsertParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var paragraphInHeaderE = document.Header.Even.InsertParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                // 2 section, 9 paragraphs + 7 pagebreaks = 15 paragraphs, 7 pagebreaks
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                // additional sections
                document.Save(openWord);
            }
        }

        private static void Example_BasicSections(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.InsertParagraph("Test 1 - Should be before 1st section").SetColor(Color.LightPink);
                
                var section1 = document.InsertSection();
                section1.InsertParagraph("Test 1 - Should be after 1st section").SetFontFamily("Tahoma").SetFontSize(20);
                
                document.InsertParagraph("Test 2 - Should be after 1st section");
                var section2 = document.InsertSection();

                document.InsertParagraph("Test 3 - Should be after 2nd section");
                document.InsertParagraph("Test 4 - Should be after 2nd section").SetBold().AppendText(" more text").SetColor(Color.DarkSalmon);

                var section3 = document.InsertSection();

                var para = document.InsertParagraph("Test 5 -");
                para = para.AppendText(" and more text");
                para.Bold = true;

                document.InsertPageBreak();

                var paragraph = document.InsertParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Blue.ToHexColor();

                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AppendText(" This is continuation").SetUnderline(UnderlineValues.Double).SetHighlight(HighlightColorValues.DarkGreen).SetFontSize(15).SetColor(Color.Aqua);

                paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Blue.ToHexColor();

                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AppendText(" This is continuation").SetUnderline(UnderlineValues.Double).SetHighlight(HighlightColorValues.DarkGreen).SetFontSize(15).SetColor(Color.Yellow);


                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 3: " + document.Sections[3].Paragraphs.Count);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 3: " + document.Sections[3].Paragraphs.Count);

                document.Save(openWord);
            }
        }

        private static void Example_BasicParagraphs(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Blue.ToHexColor();

                paragraph.AppendText(" This is continuation").SetUnderline(UnderlineValues.Double).SetFontSize(15).SetColor(Color.Yellow).SetHighlight(HighlightColorValues.DarkGreen);

                
                Console.WriteLine("+ Color: " + paragraph.Color);
                Console.WriteLine("+ Color 0: " + document.Paragraphs[0].Color);
                Console.WriteLine("+ Color 1: " + document.Paragraphs[1].Color);
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("+ Color 0: " + document.Paragraphs[0].Color);
                Console.WriteLine("+ Color 1: " + document.Paragraphs[1].Color);
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                document.Save(openWord);
            }
        }

        private static void Example_BasicWordWithHeaderAndFooter(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;

                var paragraphInHeader = document.Header.Default.InsertParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";
                
                document.InsertPageBreak();

                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var section2 = document.InsertSection();
                section2.AddHeadersAndFooters();
                
                var paragraghInHeaderSection1 = section2.Header.Default.InsertParagraph();
                paragraghInHeaderSection1.Text = "Weird shit? 1";
                
                paragraph = document.InsertParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var section3 = document.InsertSection();
                section3.AddHeadersAndFooters();

                var paragraghInHeaderSection3 = section3.Header.Default.InsertParagraph();
                paragraghInHeaderSection3.Text = "Weird shit? 2";

                paragraph = document.InsertParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

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
        
        private static void Example_BasicWordWithSections(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.InsertParagraph("Test 1");
                var section1 = document.InsertSection(SectionMarkValues.NextPage);

                document.InsertParagraph("Test 2");
                var section2 = document.InsertSection(SectionMarkValues.Continuous);

                document.InsertParagraph("Test 3");
                var section3 = document.InsertSection(SectionMarkValues.NextPage);
                section3.InsertParagraph("Paragraph added to section number 3");
                section3.InsertParagraph("Continue adding paragraphs to section 3");

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
                document.Paragraphs[1].Color = "7178a8";

                var paragraph = section1.InsertParagraph("We missed paragraph on 1 section (2nd page)");
                var newParagraph = paragraph.InsertParagraphAfterSelf();
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

        private static void Example_Load(bool openWord = false) {
            string folderPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithSection.docx");
            //filePath = @"C:\Support\GitHub\OfficeIMO\OfficeIMO.Tests\Documents\DocumentWithSection.docx";
            //filePath = @"C:\Support\GitHub\OfficeIMO\OfficeIMO.Tests\Documents\EmptyDocumentWithSection.docx";

            using (WordDocument document = WordDocument.Load(filePath, true)) {
                Console.WriteLine("+ Document Path: " + document.FilePath);
                Console.WriteLine("+ Document Title: " + document.BuiltinDocumentProperties.Title);
                Console.WriteLine("+ Document Author: " + document.BuiltinDocumentProperties.Creator);

                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);
                
                document.Open(openWord);
            }
        }

        private static void Example_LoadDocumentWithProperties(bool openWord = false) {
            string folderPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithBuiltinAndCustomProperties.docx");

            using (WordDocument document = WordDocument.Load(filePath, true)) {
                Console.WriteLine("+ Document Path: " + document.FilePath);
                Console.WriteLine("+ Document Title: " + document.BuiltinDocumentProperties.Title);
                Console.WriteLine("+ Document Author: " + document.BuiltinDocumentProperties.Creator);

                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                Console.WriteLine(document.ApplicationProperties.ApplicationVersion);

                document.Open(openWord);
            }
        }
        private static void Example_BasicCustomProperties(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                
                document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
                document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
                document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

                Console.WriteLine("+ Custom properties: " + document.CustomDocumentProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomDocumentProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomDocumentProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomDocumentProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomDocumentProperties.Keys.Count());
                
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath, false)) {
                Console.WriteLine("* Loading document...");
                Console.WriteLine("+ Custom properties: " + document.CustomDocumentProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomDocumentProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomDocumentProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomDocumentProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomDocumentProperties.Keys.GetEnumerator());

                document.CustomDocumentProperties["MyName"].Value = "Przemysław Kłys";

                document.Save(openWord);
            }
        }

        private static void Example_ValidateDocument(string filePath) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
                document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
                document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

                Console.WriteLine("+ Custom properties: " + document.CustomDocumentProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomDocumentProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomDocumentProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomDocumentProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomDocumentProperties.Keys.Count());

                document.Save();
            }
            Validation.ValidateWordDocument(filePath);
        }

        private static void Example_ValidateDocument_BeforeSave() {
            using (WordDocument document = WordDocument.Create()) {
                var paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
                document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
                document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

                Console.WriteLine("+ Custom properties: " + document.CustomDocumentProperties.Count);
                Console.WriteLine("++ TestProperty: " + document.CustomDocumentProperties["TestProperty"].Value);
                Console.WriteLine("++ MyName: " + document.CustomDocumentProperties["MyName"].Value);
                Console.WriteLine("++ IsTodayGreatDay: " + document.CustomDocumentProperties["IsTodayGreatDay"].Value);
                Console.WriteLine("++ Count: " + document.CustomDocumentProperties.Keys.Count());

                document.ValidateDocument();
            }
        }

        private static void Example_BasicLists(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList = document.AddList(ListStyles.Headings111);
                wordList.AddItem("Text 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList.AddItem("Text 2.1", 1).SetColor(Color.Brown);
                wordList.AddItem("Text 2.2", 1).SetColor(Color.Brown);
                wordList.AddItem("Text 2.3", 1).SetColor(Color.Brown);
                wordList.AddItem("Text 2.3.4", 2).SetColor(Color.Brown);
                // here we set another list element but we also change it using standard paragraph change
                paragraph = wordList.AddItem("Text 3");
                paragraph.Bold = true;
                paragraph.SetItalic();

                paragraph = document.InsertParagraph("This is second list").SetColor(Color.OrangeRed).SetUnderline(UnderlineValues.Double);
                
                WordList wordList1 = document.AddList(ListStyles.HeadingIA1);
                wordList1.AddItem("Temp 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList1.AddItem("Temp 2.1", 1).SetColor(Color.Brown);
                wordList1.AddItem("Temp 2.2", 1).SetColor(Color.Brown);
                wordList1.AddItem("Temp 2.3", 1).SetColor(Color.Brown);
                wordList1.AddItem("Temp 2.3.4", 2).SetColor(Color.Brown).Remove();
                wordList1.ListItems[1].Remove();
                paragraph = wordList1.AddItem("Temp 3");

                paragraph = document.InsertParagraph("This is third list").SetColor(Color.Blue).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(ListStyles.BulletedChars);
                wordList2.AddItem("Text 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList2.AddItem("Text 2.1", 1).SetColor(Color.Brown);
                wordList2.AddItem("Text 2.2", 1).SetColor(Color.Brown);
                wordList2.AddItem("Text 2.3", 1).SetColor(Color.Brown);
                wordList2.AddItem("Text 2.3.4", 2).SetColor(Color.Brown);

                paragraph = document.InsertParagraph("This is fourth list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList3 = document.AddList(ListStyles.Heading1ai);
                wordList3.AddItem("Text 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList3.AddItem("Text 2.1", 1).SetColor(Color.Brown);
                wordList3.AddItem("Text 2.2", 1).SetColor(Color.Brown);
                wordList3.AddItem("Text 2.3", 1).SetColor(Color.Brown);
                wordList3.AddItem("Text 2.3.4", 2).SetColor(Color.Brown);

                paragraph = document.InsertParagraph("This is five list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList4 = document.AddList(ListStyles.Headings111Shifted);
                wordList4.AddItem("Text 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList4.AddItem("Text 2.1", 1).SetColor(Color.Brown);
                wordList4.AddItem("Text 2.2", 1).SetColor(Color.Brown);
                wordList4.AddItem("Text 2.3", 1).SetColor(Color.Brown);
                wordList4.AddItem("Text 2.3.4", 2).SetColor(Color.Brown);
                
                document.Save(openWord);
            }
        }


        private static void Example_BasicLists2(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList1 = document.AddList(ListStyles.ArticleSections);
                wordList1.AddItem("Text 1");
                wordList1.AddItem("Text 2", 1);
                wordList1.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is second list").SetColor(Color.OrangeRed).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(ListStyles.Headings111);
                wordList2.AddItem("Temp 2");
                wordList2.AddItem("Text 2", 1);
                wordList2.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is third list").SetColor(Color.Blue).SetUnderline(UnderlineValues.Double);

                WordList wordList3 = document.AddList(ListStyles.HeadingIA1);
                wordList3.AddItem("Text 3");
                wordList3.AddItem("Text 2", 1);
                wordList3.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is fourth list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList4 = document.AddList(ListStyles.Chapters); // Chapters support only level 0
                wordList4.AddItem("Text 1");
                wordList4.AddItem("Text 2");
                wordList4.AddItem("Text 3");
                
                paragraph = document.InsertParagraph("This is five list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList5 = document.AddList(ListStyles.BulletedChars);
                wordList5.AddItem("Text 5");
                wordList5.AddItem("Text 2", 1);
                wordList5.AddItem("Text 3", 2);
                
                paragraph = document.InsertParagraph("This is 6th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);
                
                WordList wordList6 = document.AddList(ListStyles.Heading1ai);
                wordList6.AddItem("Text 6");
                wordList6.AddItem("Text 2", 1);
                wordList6.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is 7th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);
                
                WordList wordList7 = document.AddList(ListStyles.Headings111Shifted);
                wordList7.AddItem("Text 7");
                wordList7.AddItem("Text 2", 1);
                wordList7.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is 7th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);
                
                WordList wordList8 = document.AddList(ListStyles.Bulleted);
                wordList8.AddItem("Text 8");
                wordList8.AddItem("Text 8.1", 1);
                wordList8.AddItem("Text 8.2", 2);

                document.Save(openWord);
            }
        }
    }
}