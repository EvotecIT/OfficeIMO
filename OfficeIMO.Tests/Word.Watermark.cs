using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithWatermark() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                document.Sections[0].SetMargins(WordMargin.Normal);

                document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Watermark");

                document.AddSection();
                document.Sections[1].AddHeadersAndFooters();
                document.Sections[1].Margins.Type = WordMargin.Narrow;

               
                document.Sections[1].AddWatermark(WordWatermarkStyle.Text, "Draft");

                document.Settings.SetBackgroundColor(Color.Azure);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx"))) {
               
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx"))) {
                document.Save();
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithWatermark2() {
            // this test adding watermark without adding headers/footers first (watermark is added in the header.default)
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Watermark");
                document.AddSection();
                document.Sections[1].AddWatermark(WordWatermarkStyle.Text, "Draft");
                document.Settings.SetBackgroundColor(Color.Azure);
                document.AddSection();
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx"))) {

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx"))) {
                document.Save();
            }
        }
    }
}
