using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = System.Drawing.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_OpeningWordWithSections() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "BasicDocumentWithSections.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 3);

                // There is only one PageBreak in this document.
                Assert.True(document.Sections.Count == 4);

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
        [Fact]
        public void Test_OpeningWordEmptyDocument() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "EmptyDocument.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 1);

                // There is only one PageBreak in this document.
                Assert.True(document.Sections.Count == 1);

                Assert.True(document.Paragraphs[0].IsEmpty == true, "Paragraph is not empty");

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
        [Fact]
        public void Test_OpeningWordEmptyDocumentWithSectionBreak() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "EmptyDocumentWithSection.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs is wrong.");

                Assert.True(document.Paragraphs[0].IsEmpty == true, "Paragraph is not empty");

                // There is only one PageBreak in this document.
                Assert.True(document.Sections.Count == 2, "Number of sections is wrong.");

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
        [Fact]
        public void Test_OpeningWordDocumentWithSectionBreak() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "DocumentWithSection.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 2, "Number of paragraphs is wrong.");
                Assert.True(document.Sections.Count == 2, "Number of sections is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs in first section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[1].Paragraphs.Count == 1, "Number of paragraphs in second section is wrong. Current: " + document.Sections[1].Paragraphs.Count);
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentWithSectionBreak() {
            //using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithSection.docx"))) {
            //    // There is only one Paragraph at the document level.
            //    Assert.True(document.Paragraphs.Count == 2, "Number of paragraphs is wrong.");

            //    // There is only one PageBreak in this document.
            //    Assert.True(document.Sections.Count == 2, "Number of sections is wrong.");

            //    Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of sections is wrong.");
            //    Assert.True(document.Sections[1].Paragraphs.Count == 1, "Number of sections is wrong.");
            //}
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithSections.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test 1");

                Assert.True(document.Paragraphs[0].IsEmpty == false, "Paragraph is empty");

                var section1 = document.AddSection();
                section1.AddParagraph("Test 1");

                document.AddParagraph("Test 2");
                var section2 = document.AddSection();

                document.AddParagraph("Test 3");
                var section3 = document.AddSection();

                Assert.True(document.Paragraphs.Count == 4, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 2, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 0, "Number of paragraphs on 4th section is wrong.");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSections.docx"))) {
                // There is only one Paragraph at the document level.

                Assert.True(document.Paragraphs[0].IsEmpty == false, "Paragraph is empty");

                Assert.True(document.Paragraphs.Count == 4, "Number of paragraphs during load is wrong.");
                Assert.True(document.Sections.Count == 4, "Number of sections during load is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 2, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 0, "Number of paragraphs on 4th section is wrong.");
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentWithSectionsPageOrientation() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithSections1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section1").SetColor(Color.LightPink);

                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;

                section1.AddParagraph("Test Section2").SetFontFamily("Tahoma").SetFontSize(20);

                section1.ColumnCount = 2;
                for (int i = 0; i < 50; i++) {
                    section1.AddParagraph("Test Section2 - Multicolumn");
                }

                var section2 = document.AddSection();

                section2.AddParagraph("Test Section3").SetFontFamily("Tahoma").SetFontSize(20);

                section2.PageOrientation = PageOrientationValues.Landscape;

                Assert.True(document.Paragraphs.Count == 53, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 51, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");

                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[1].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Sections[2].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");

                Assert.True(document.Sections[0].ColumnCount == null, "Columns count for section should match");
                Assert.True(document.Sections[1].ColumnCount == 2, "Columns count for section should match");
                Assert.True(document.Sections[2].ColumnCount == 2, "Columns count for section should match");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSections1.docx"))) {
                Assert.True(document.Paragraphs.Count == 53, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 51, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");

                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[1].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Sections[2].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");

                Assert.True(document.Sections[0].ColumnCount == null, "Columns count for section should match");
                Assert.True(document.Sections[1].ColumnCount == 2, "Columns count for section should match");
                Assert.True(document.Sections[2].ColumnCount == 2, "Columns count for section should match");


                var section1 = document.AddSection();
                section1.AddParagraph("Test Section4");
                // when adding section column count, page orientation is copied from section before
                // we reset it to one
                section1.ColumnCount = 1;

                var section2 = document.AddSection();
                section2.AddParagraph("Test Section5");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Section6");
                section3.PageOrientation = PageOrientationValues.Portrait;

                Assert.True(document.Paragraphs.Count == 56, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 6, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 51, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[4].Paragraphs.Count == 1, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[5].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");

                Assert.True(document.Sections[0].Paragraphs[0].Text == "Test Section1", "Paragraph text must match.");
                Assert.True(document.Sections[1].Paragraphs[0].Text == "Test Section2", "Paragraph text must match.");
                Assert.True(document.Sections[2].Paragraphs[0].Text == "Test Section3", "Paragraph text must match.");
                Assert.True(document.Sections[3].Paragraphs[0].Text == "Test Section4", "Paragraph text must match.");
                Assert.True(document.Sections[4].Paragraphs[0].Text == "Test Section5", "Paragraph text must match.");
                Assert.True(document.Sections[5].Paragraphs[0].Text == "Test Section6", "Paragraph text must match.");

                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[1].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Sections[2].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[3].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[4].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[5].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");

                document.AddParagraph("This goes to last section 1");
                document.AddParagraph("This goes to last section 2");

                section1.AddParagraph("This goes to section 4");

                Assert.True(document.Sections[5].Paragraphs[1].Text == "This goes to last section 1", "Paragraph text must match.");
                Assert.True(document.Sections[5].Paragraphs[2].Text == "This goes to last section 2", "Paragraph text must match.");
                Assert.True(document.Sections[3].Paragraphs[1].Text == "This goes to section 4", "Paragraph text must match.");

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSections1.docx"))) {

                Assert.True(document.Paragraphs.Count == 59, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 6, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 51, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 2, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[4].Paragraphs.Count == 1, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[5].Paragraphs.Count == 3, "Number of paragraphs on 3rd section is wrong.");

                Assert.True(document.Sections[0].Paragraphs[0].Text == "Test Section1", "Paragraph text must match.");
                Assert.True(document.Sections[1].Paragraphs[0].Text == "Test Section2", "Paragraph text must match.");
                Assert.True(document.Sections[2].Paragraphs[0].Text == "Test Section3", "Paragraph text must match.");
                Assert.True(document.Sections[3].Paragraphs[0].Text == "Test Section4", "Paragraph text must match.");
                Assert.True(document.Sections[4].Paragraphs[0].Text == "Test Section5", "Paragraph text must match.");
                Assert.True(document.Sections[5].Paragraphs[0].Text == "Test Section6", "Paragraph text must match.");

                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[1].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Sections[2].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[3].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[4].PageOrientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Sections[5].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");

                Assert.True(document.Sections[5].Paragraphs[1].Text == "This goes to last section 1", "Paragraph text must match.");
                Assert.True(document.Sections[5].Paragraphs[2].Text == "This goes to last section 2", "Paragraph text must match.");
                Assert.True(document.Sections[3].Paragraphs[1].Text == "This goes to section 4", "Paragraph text must match.");

                Assert.True(document.Sections[0].ColumnCount == null, "Columns count for section should match");
                Assert.True(document.Sections[1].ColumnCount == 2, "Columns count for section should match");
                Assert.True(document.Sections[2].ColumnCount == 2, "Columns count for section should match");
                Assert.True(document.Sections[3].ColumnCount == 1, "Columns count for section should match");
                Assert.True(document.Sections[4].ColumnCount == 1, "Columns count for section should match");
                Assert.True(document.Sections[5].ColumnCount == 1, "Columns count for section should match");
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentWithPageMargins() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageMargins.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                document.AddParagraph("Section 0");
                document.Sections[0].SetMargins(PageMargins.Normal);

                document.AddSection();
                document.Sections[1].SetMargins(PageMargins.Narrow);
                document.AddParagraph("Section 1");

                document.AddSection();
                document.Sections[2].SetMargins(PageMargins.Mirrored);
                document.AddParagraph("Section 2");

                document.AddSection();
                document.Sections[3].SetMargins(PageMargins.Moderate);
                document.AddParagraph("Section 3");

                document.AddSection();
                document.Sections[4].SetMargins(PageMargins.Wide);
                document.AddParagraph("Section 4");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 5, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 1, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 1, "Number of paragraphs on 4th section is wrong.");
                Assert.True(document.Sections[4].Paragraphs.Count == 1, "Number of paragraphs on 5th section is wrong.");

                // Normal
                Assert.True(document.Sections[0].Margins.Left == 1440);
                Assert.True(document.Sections[0].Margins.Right == 1440);
                Assert.True(document.Sections[0].Margins.Top == 1440);
                Assert.True(document.Sections[0].Margins.Bottom == 1440);
                Assert.True(document.Sections[0].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[0].Margins.FooterDistance == 720);

                //Narrow
                Assert.True(document.Sections[1].Margins.Left == 720);
                Assert.True(document.Sections[1].Margins.Right == 720);
                Assert.True(document.Sections[1].Margins.Top == 720);
                Assert.True(document.Sections[1].Margins.Bottom == 720);
                Assert.True(document.Sections[1].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[1].Margins.FooterDistance == 720);

                // Mirrored
                Assert.True(document.Sections[2].Margins.Left == 1800);
                Assert.True(document.Sections[2].Margins.Right == 1440);
                Assert.True(document.Sections[2].Margins.Top == 1440);
                Assert.True(document.Sections[2].Margins.Bottom == 1440);
                Assert.True(document.Sections[2].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[2].Margins.FooterDistance == 720);

                // Moderate
                Assert.True(document.Sections[3].Margins.Left == 1080);
                Assert.True(document.Sections[3].Margins.Right == 1080);
                Assert.True(document.Sections[3].Margins.Top == 1440);
                Assert.True(document.Sections[3].Margins.Bottom == 1440);
                Assert.True(document.Sections[3].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[3].Margins.FooterDistance == 720);

                // Wide
                Assert.True(document.Sections[4].Margins.Left == 2880);
                Assert.True(document.Sections[4].Margins.Right == 2880);
                Assert.True(document.Sections[4].Margins.Top == 1440);
                Assert.True(document.Sections[4].Margins.Bottom == 1440);
                Assert.True(document.Sections[4].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[4].Margins.FooterDistance == 720);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageMargins.docx"))) {
                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 5, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 1, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 1, "Number of paragraphs on 4th section is wrong.");
                Assert.True(document.Sections[4].Paragraphs.Count == 1, "Number of paragraphs on 5th section is wrong.");

                // Normal
                Assert.True(document.Sections[0].Margins.Left == 1440);
                Assert.True(document.Sections[0].Margins.Right == 1440);
                Assert.True(document.Sections[0].Margins.Top == 1440);
                Assert.True(document.Sections[0].Margins.Bottom == 1440);
                Assert.True(document.Sections[0].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[0].Margins.FooterDistance == 720);

                //Narrow
                Assert.True(document.Sections[1].Margins.Left == 720);
                Assert.True(document.Sections[1].Margins.Right == 720);
                Assert.True(document.Sections[1].Margins.Top == 720);
                Assert.True(document.Sections[1].Margins.Bottom == 720);
                Assert.True(document.Sections[1].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[1].Margins.FooterDistance == 720);

                // Mirrored
                Assert.True(document.Sections[2].Margins.Left == 1800);
                Assert.True(document.Sections[2].Margins.Right == 1440);
                Assert.True(document.Sections[2].Margins.Top == 1440);
                Assert.True(document.Sections[2].Margins.Bottom == 1440);
                Assert.True(document.Sections[2].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[2].Margins.FooterDistance == 720);

                // Moderate
                Assert.True(document.Sections[3].Margins.Left == 1080);
                Assert.True(document.Sections[3].Margins.Right == 1080);
                Assert.True(document.Sections[3].Margins.Top == 1440);
                Assert.True(document.Sections[3].Margins.Bottom == 1440);
                Assert.True(document.Sections[3].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[3].Margins.FooterDistance == 720);

                // Wide
                Assert.True(document.Sections[4].Margins.Left == 2880);
                Assert.True(document.Sections[4].Margins.Right == 2880);
                Assert.True(document.Sections[4].Margins.Top == 1440);
                Assert.True(document.Sections[4].Margins.Bottom == 1440);
                Assert.True(document.Sections[4].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[4].Margins.FooterDistance == 720);

                document.AddSection();
                document.AddParagraph("Section 5");
                document.Sections[5].SetMargins(PageMargins.Normal);


                document.AddSection();
                document.AddParagraph("Section 6");
                document.Sections[6].SetMargins(PageMargins.Office2003Default);


                // Normal
                Assert.True(document.Sections[5].Margins.Left == 1440);
                Assert.True(document.Sections[5].Margins.Right == 1440);
                Assert.True(document.Sections[5].Margins.Top == 1440);
                Assert.True(document.Sections[5].Margins.Bottom == 1440);
                Assert.True(document.Sections[5].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[5].Margins.FooterDistance == 720);

                // Office2003Default
                Assert.True(document.Sections[6].Margins.Left == 1800);
                Assert.True(document.Sections[6].Margins.Right == 1800);
                Assert.True(document.Sections[6].Margins.Top == 1440);
                Assert.True(document.Sections[6].Margins.Bottom == 1440);
                Assert.True(document.Sections[6].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[6].Margins.FooterDistance == 720);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageMargins.docx"))) {
                Assert.True(document.Paragraphs.Count == 7, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 7, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 1, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 1, "Number of paragraphs on 4th section is wrong.");
                Assert.True(document.Sections[4].Paragraphs.Count == 1, "Number of paragraphs on 5th section is wrong.");
                Assert.True(document.Sections[5].Paragraphs.Count == 1, "Number of paragraphs on 6th section is wrong.");
                Assert.True(document.Sections[6].Paragraphs.Count == 1, "Number of paragraphs on 6th section is wrong.");

                // Normal
                Assert.True(document.Sections[0].Margins.Left == 1440);
                Assert.True(document.Sections[0].Margins.Right == 1440);
                Assert.True(document.Sections[0].Margins.Top == 1440);
                Assert.True(document.Sections[0].Margins.Bottom == 1440);
                Assert.True(document.Sections[0].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[0].Margins.FooterDistance == 720);

                //Narrow
                Assert.True(document.Sections[1].Margins.Left == 720);
                Assert.True(document.Sections[1].Margins.Right == 720);
                Assert.True(document.Sections[1].Margins.Top == 720);
                Assert.True(document.Sections[1].Margins.Bottom == 720);
                Assert.True(document.Sections[1].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[1].Margins.FooterDistance == 720);

                // Mirrored
                Assert.True(document.Sections[2].Margins.Left == 1800);
                Assert.True(document.Sections[2].Margins.Right == 1440);
                Assert.True(document.Sections[2].Margins.Top == 1440);
                Assert.True(document.Sections[2].Margins.Bottom == 1440);
                Assert.True(document.Sections[2].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[2].Margins.FooterDistance == 720);

                // Moderate
                Assert.True(document.Sections[3].Margins.Left == 1080);
                Assert.True(document.Sections[3].Margins.Right == 1080);
                Assert.True(document.Sections[3].Margins.Top == 1440);
                Assert.True(document.Sections[3].Margins.Bottom == 1440);
                Assert.True(document.Sections[3].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[3].Margins.FooterDistance == 720);

                // Wide
                Assert.True(document.Sections[4].Margins.Left == 2880);
                Assert.True(document.Sections[4].Margins.Right == 2880);
                Assert.True(document.Sections[4].Margins.Top == 1440);
                Assert.True(document.Sections[4].Margins.Bottom == 1440);
                Assert.True(document.Sections[4].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[4].Margins.FooterDistance == 720);

                // Normal
                Assert.True(document.Sections[5].Margins.Left == 1440);
                Assert.True(document.Sections[5].Margins.Right == 1440);
                Assert.True(document.Sections[5].Margins.Top == 1440);
                Assert.True(document.Sections[5].Margins.Bottom == 1440);
                Assert.True(document.Sections[5].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[5].Margins.FooterDistance == 720);

                // Office2003Default
                Assert.True(document.Sections[6].Margins.Left == 1800);
                Assert.True(document.Sections[6].Margins.Right == 1800);
                Assert.True(document.Sections[6].Margins.Top == 1440);
                Assert.True(document.Sections[6].Margins.Bottom == 1440);
                Assert.True(document.Sections[6].Margins.HeaderDistance == 720);
                Assert.True(document.Sections[6].Margins.FooterDistance == 720);
            }
        }
    }
}
