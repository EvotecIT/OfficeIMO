using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AcceptRevisions_RemovesTrackedChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChanges.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Before");

                var paragraph = document.AddParagraph();

                var ins = new InsertedRun() { Author = "Codex", Date = DateTime.Now };
                ins.AppendChild(new Run(new Text("Added")));
                paragraph._paragraph.Append(ins);

                var del = new DeletedRun() { Author = "Codex", Date = DateTime.Now };
                del.AppendChild(new Run(new Text("Removed")));
                paragraph._paragraph.Append(del);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document._document.Body.Descendants<InsertedRun>().Any());
                Assert.True(document._document.Body.Descendants<DeletedRun>().Any());

                document.AcceptRevisions();

                Assert.False(document._document.Body.Descendants<InsertedRun>().Any());
                Assert.False(document._document.Body.Descendants<DeletedRun>().Any());
                Assert.Equal(2, document.Paragraphs.Count);
                Assert.Equal("Before", document.Paragraphs[0].Text);
                Assert.Equal("Added", document.Paragraphs[1].Text);
            }
        }
    }
}
