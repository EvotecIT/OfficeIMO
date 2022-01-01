using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_OpeningWordWithSections() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "BasicDocumentWithSections.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 3);

                // There is only one PageBreak in this document.
                Assert.True(document.Sections.Count == 3);

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
    }
}
