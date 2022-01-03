using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordDocumentWithPageSize() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithPageSize.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
            
                Assert.True(document.PageOrientation == PageOrientationValues.Portrait, "Starting page orientation should be portrait");

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;

                Assert.True(document.PageOrientation == PageOrientationValues.Landscape, "Middle page orientation should be landscape when using section 0");

                document.PageOrientation = PageOrientationValues.Portrait;

                Assert.True(document.PageOrientation == PageOrientationValues.Portrait, "Middle page orientation should be portrait when using document");
                
                document.InsertParagraph("Test");

                document.PageOrientation = PageOrientationValues.Landscape;
                Assert.True(document.PageOrientation == PageOrientationValues.Landscape, "End page orientation should be landscape when using document");
                
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong.");
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithPageSize.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during load is wrong.");
                Assert.True(document.Sections.Count == 1, "Number of sections during load is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.PageOrientation == PageOrientationValues.Landscape, "Page orientation should be landscape when using document");
                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Landscape, "Page orientation should be landscape when using sections");
            }
        }
    }
}
