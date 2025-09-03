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
    }
}
