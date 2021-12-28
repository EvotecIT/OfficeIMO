using System;
using System.ComponentModel;
using System.IO;
using DocumentFormat.OpenXml.Drawing;
using Xunit;
using Path = System.IO.Path;

namespace OfficeIMO.Tests
{
    public class Word
    {
        private readonly string _directoryDocuments;
        private readonly string _directoryWithFiles;

        private static void Setup(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            } else {
                Directory.Delete(path, true);
                Directory.CreateDirectory(path);
            }
        }
        public Word() {
            _directoryDocuments = Path.Combine(Path.GetTempPath(), "DocXTests", "documents");
            Setup(_directoryDocuments); // prepare temp documents directory 
            _directoryWithFiles = TestHelper.DirectoryWithFiles;
        }


        [Fact]
        public void Test_SimpleWordDocumentCreation() {
            var filePath = Path.Combine(_directoryDocuments, "TestFile.docx");

            var path = File.Exists(filePath);
            Assert.False(path); // MUST BE FALSE

            WordDocument document = WordDocument.Create(filePath);

            document.Save();

            path = File.Exists(filePath);
            Assert.True(path);
        }

        [Fact]
        public void Test_OpeningWordAndParagraphCountMatches()
        {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "BasicDocumentWith12.docx")))
            {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 12);

                // There is only one Table in this document.
                //Assert.True(document.Tables.Count() == 1);

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
    }
}
