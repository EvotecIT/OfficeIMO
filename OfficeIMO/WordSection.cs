using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordSection {
        public WordSection(WordDocument document) {
            WordParagraph paragraph = new WordParagraph();
            WordSection section = new WordSection(document, paragraph);
        }
        public WordSection(WordDocument document, WordParagraph paragraph) {
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            
            SectionProperties sectionProperties = new SectionProperties();
            SectionType sectionType = new SectionType() {Val = SectionMarkValues.NextPage};


            sectionProperties.Append(sectionType);
            paragraphProperties.Append(sectionProperties);
            paragraph._paragraph.Append(paragraphProperties);
        }
        private static void AddSectionBreakToTheDocument(string fileName) {
            using (WordprocessingDocument mydoc = WordprocessingDocument.Open(fileName, true)) {
                MainDocumentPart myMainPart = mydoc.MainDocumentPart;
                Paragraph paragraphSectionBreak = new Paragraph();
                ParagraphProperties paragraphSectionBreakProperties = new ParagraphProperties();
                SectionProperties SectionBreakProperties = new SectionProperties();
                SectionType SectionBreakType = new SectionType() { Val = SectionMarkValues.NextPage };
                SectionBreakProperties.Append(SectionBreakType);
                paragraphSectionBreakProperties.Append(SectionBreakProperties);
                paragraphSectionBreak.Append(paragraphSectionBreakProperties);
                myMainPart.Document.Body.InsertAfter(paragraphSectionBreak, myMainPart.Document.Body.LastChild);
                myMainPart.Document.Save();
            }
        }
    }
}
