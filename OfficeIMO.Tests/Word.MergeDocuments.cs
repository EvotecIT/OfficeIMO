using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MergingDocumentsWithSeparateLists() {
            var filePath1 = Path.Combine(_directoryWithFiles, "MergeDoc1.docx");
            var filePath2 = Path.Combine(_directoryWithFiles, "MergeDoc2.docx");

            using (var doc1 = WordDocument.Create(filePath1)) {
                var list1 = doc1.AddList(WordListStyle.Headings111);
                list1.AddItem("Item 1");
                list1.AddItem("Item 2");
                doc1.Save();
            }

            using (var doc2 = WordDocument.Create(filePath2)) {
                var list2 = doc2.AddList(WordListStyle.Headings111);
                list2.AddItem("Second 1");
                list2.AddItem("Second 2");
                doc2.Save();
            }

            using (var doc1 = WordDocument.Load(filePath1))
            using (var doc2 = WordDocument.Load(filePath2)) {
                doc1.AppendDocument(doc2);
                doc1.Save();
            }

            using (var merged = WordDocument.Load(filePath1)) {
                Assert.Equal(2, merged.Lists.Count);
                var numbering = merged._wordprocessingDocument.MainDocumentPart
                    .NumberingDefinitionsPart!.Numbering;
                var ids = numbering.Elements<NumberingInstance>()
                    .Select(n => n.NumberID.Value).Distinct().ToList();
                Assert.Equal(2, ids.Count);
            }
        }
    }
}
