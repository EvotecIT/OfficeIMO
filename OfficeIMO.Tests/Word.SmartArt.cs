using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddSmartArt() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArtDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddSmartArt(SmartArtType.BasicProcess);
                var mainPart = document._wordprocessingDocument.MainDocumentPart!;
                Assert.Single(mainPart.DiagramDataParts);
                Assert.Single(mainPart.DiagramLayoutDefinitionParts);
                Assert.Single(mainPart.DiagramStyleParts);
                Assert.Single(mainPart.DiagramColorsParts);
                Assert.Single(document.SmartArts);
                Assert.Single(document.Sections[0].SmartArts);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var mainPart = document._wordprocessingDocument.MainDocumentPart!;
                Assert.Single(mainPart.DiagramDataParts);
                Assert.Single(mainPart.DiagramLayoutDefinitionParts);
                Assert.Single(mainPart.DiagramStyleParts);
                Assert.Single(mainPart.DiagramColorsParts);
                Assert.Single(document.SmartArts);
                Assert.Single(document.Sections[0].SmartArts);
            }
        }

        [Fact]
        public void Test_SmartArt_Retrieval_After_Load() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArtRetrieve.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddSmartArt(SmartArtType.Hierarchy);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.SmartArts);
                Assert.Single(document.Sections[0].SmartArts);
                Assert.Single(document.ParagraphsSmartArts);
            }
        }
    }
}
