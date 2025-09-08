using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentSectionBuilderOrientation() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentSectionBuilderOrientation.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Section(s => s.New().Orientation(PageOrientationValues.Landscape))
                    .End()
                    .Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(PageOrientationValues.Landscape, document.Sections[1].PageSettings.Orientation);
            }
        }

        [Fact]
        public void Test_FluentSectionBuilderBackground() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentSectionBuilderBackground.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Section(s => s.New().Background("FF00FF"))
                    .End()
                    .Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("ff00ff", document.Background.Color);
            }
        }

        [Fact]
        public void Test_FluentSectionBuilderContentAssignments() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentSectionBuilderContent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Section1"))
                    .Section(s => s.New()
                        .Paragraph(p => p.Text("Section2"))
                        .Table(t => t.Row("cell")))
                    .End()
                    .Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Sections.Count);
                Assert.Single(document.Sections[0].Paragraphs);
                Assert.Empty(document.Sections[0].Tables);
                Assert.Single(document.Sections[1].Paragraphs);
                Assert.Single(document.Sections[1].Tables);
                Assert.Equal("Section2", document.Sections[1].Paragraphs[0].Text);
            }
        }
    }
}
