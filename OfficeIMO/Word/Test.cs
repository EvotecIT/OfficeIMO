using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal class Test {
        public static void OpenAndAddTextToWordDocument(string filepath, string txt) {
            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true);
            MainDocumentPart part = wordprocessingDocument.MainDocumentPart;
            Body body = part.Document.Body;
            //create a new footer Id=rIdf2
            FooterPart footerPart2 = part.AddNewPart<FooterPart>("rIdf2");
            GenerateFooterPartContent(footerPart2);
            //create a new header Id=rIdh2
            HeaderPart headerPart2 = part.AddNewPart<HeaderPart>("rIdh2");
            GenerateHeaderPartContent(headerPart2);
            //replace the attribute of SectionProperties to add new footer and header
            SectionProperties lxml = body.GetFirstChild<SectionProperties>();
            lxml.GetFirstChild<HeaderReference>().Remove();
            lxml.GetFirstChild<FooterReference>().Remove();
            HeaderReference headerReference1 = new HeaderReference() {
                Type = HeaderFooterValues.Default, Id = "rIdh2"
            };
            FooterReference footerReference1 = new FooterReference() {
                Type = HeaderFooterValues.Default, Id = "rIdf2"
            };
            lxml.Append(headerReference1);
            lxml.Append(footerReference1);
            //add the correlation of last Paragraph
            OpenXmlElement oxl = body.ChildElements.GetItem(body.ChildElements.Count - 2);
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = oxl.GetAttribute("rsidR", oxl.NamespaceUri).Value };
            HeaderReference headerReference2 = new HeaderReference() {
                Type = HeaderFooterValues.Default, Id = part.GetIdOfPart(part.HeaderParts.FirstOrDefault())
            };
            FooterReference footerReference2 = new FooterReference() {
                Type = HeaderFooterValues.Default, Id = part.GetIdOfPart(part.FooterParts.FirstOrDefault())
            };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };
            sectionProperties1.Append(headerReference2);
            sectionProperties1.Append(footerReference2);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);
            paragraphProperties1.Append(sectionProperties1);
            oxl.InsertAt<ParagraphProperties>(paragraphProperties1, 0);
            body.InsertBefore<Paragraph>(GenerateParagraph(txt, oxl.GetAttribute("rsidRDefault", oxl.NamespaceUri).Value), body.GetFirstChild<SectionProperties>());
            part.Document.Save();
            wordprocessingDocument.Close();
        }

        //Generate new Paragraph
        public static Paragraph GenerateParagraph(string text, string rsidR) {
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = rsidR };
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Left, Position = 5583 };
            tabs1.Append(tabStop1);
            paragraphProperties1.Append(tabs1);
            Run run1 = new Run();
            //Text text1 = new Text();
            //text1.Text = text;
            //run1.Append(text1);
            Run run2 = new Run();
            TabChar tabChar1 = new TabChar();
            run2.Append(tabChar1);
            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            return paragraph1;
        }

        static void GenerateHeaderPartContent(HeaderPart hpart) {
            Header header1 = new Header();
            Paragraph paragraph1 = new Paragraph();
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };
            paragraphProperties1.Append(paragraphStyleId1);
            Run run1 = new Run();
            //Text text1 = new Text();
            //text1.Text = "";
            //run1.Append(text1);
            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            header1.Append(paragraph1);
            hpart.Header = header1;
        }

        static void GenerateFooterPartContent(FooterPart fpart) {
            Footer footer1 = new Footer();
            Paragraph paragraph1 = new Paragraph();
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };
            paragraphProperties1.Append(paragraphStyleId1);
            Run run1 = new Run();
            //Text text1 = new Text();
            //text1.Text = "";
            //run1.Append(text1);
            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            footer1.Append(paragraph1);
            fpart.Footer = footer1;
        }
    }
}