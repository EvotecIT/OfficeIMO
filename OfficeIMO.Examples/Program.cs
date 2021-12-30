using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Helper;

namespace OfficeIMO.Examples {
    internal class Program {
        private static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            } else {
                Directory.Delete(path, true);
                Directory.CreateDirectory(path);
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

            Console.WriteLine("[*] Creating standard document with paragraph");
            filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithParagraphs.docx");
            Example_BasicWord(filePath, true);

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
            //Example_SectionsToDocument(filePath, true);
        }

        private static void Example_BasicEmptyWord(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Title = "This is my title";
                document.Creator = "Przemysław Kłys";
                document.Keywords = "word, docx, test";
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

        private static void Example_BasicDocumentProperties(string filePath,bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Title = "This is my title";
                document.Creator = "Przemysław Kłys";
                document.Keywords = "word, docx, test";

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
            document.Title = "This is sparta";
            document.Creator = "Przemek";

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
            Console.WriteLine("+ Document Title: " + document.Title);
            Console.WriteLine("+ Document Author: " + document.Creator);

            document.Dispose();
        }
        private static void Example_ReadWordWithImages() {
            string outputPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "BasicDocumentWithImages.docx"), true);
            Console.WriteLine("+ Document paragraphs: " + document.Paragraphs.Count);
            Console.WriteLine("+ Document images: " + document.Images.Count);
            
            document.Images[0].SaveToFile(System.IO.Path.Combine(outputPath,"random.jpg"));
        }
        private static void Example_SectionsToDocument(string filePath, bool openWord) {
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Title = "This is my title";
                document.Creator = "Przemysław Kłys";
                document.Keywords = "word, docx, test";

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
                //paragraph1.AppendText("Text4 - Which we will remove");

                //document.Paragraphs.Last().Remove();

                paragraph = document.InsertParagraph("Last paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AppendText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;

                document.Save(openWord);

            }
        }
    }
}