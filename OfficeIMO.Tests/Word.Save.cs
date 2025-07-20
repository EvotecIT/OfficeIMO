using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains tests for saving Word documents.
    /// </summary>
    public partial class Word {

        [Fact]
        public void Test_Save() {
            var filePath1 = Path.Combine(_directoryWithFiles, "FirstDocument11.docx");
            var filePath2 = Path.Combine(_directoryWithFiles, "FirstDocument12.docx");
            var filePath3 = Path.Combine(_directoryWithFiles, "FirstDocument13.docx");
            var filePath4 = Path.Combine(_directoryWithFiles, "FirstDocument14.docx");

            File.Delete(filePath1);
            File.Delete(filePath2);
            File.Delete(filePath3);
            File.Delete(filePath4);

            using (var document = WordDocument.Create()) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";

                Assert.Empty(document.Paragraphs);

                document.AddParagraph("This is my test in document 1");

                Assert.False(File.Exists(filePath1));

                document.Save(filePath1);

                Assert.True(File.Exists(filePath1));
                Assert.False(filePath1.IsFileLocked());

                Assert.Single(document.Paragraphs);

                document.AddParagraph("This is my test in document 2");

                document.Save(filePath2);

                Assert.Equal(2, document.Paragraphs.Count);

                Assert.True(File.Exists(filePath2));
                Assert.False(filePath2.IsFileLocked());

                document.AddParagraph("This is my test in document 3");

                Assert.Equal(3, document.Paragraphs.Count);

                document.Save(filePath3);

                Assert.True(File.Exists(filePath3));
                Assert.False(filePath3.IsFileLocked());

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }

            Assert.False(filePath1.IsFileLocked());
            Assert.False(filePath2.IsFileLocked());
            Assert.False(filePath3.IsFileLocked());

            using (var document = WordDocument.Load(filePath1)) {
                Assert.Single(document.Paragraphs);
                Assert.False(filePath1.IsFileLocked());
            }
            using (var document = WordDocument.Load(filePath2)) {
                Assert.Equal(2, document.Paragraphs.Count);
                Assert.False(filePath2.IsFileLocked());
            }
            using (var document = WordDocument.Load(filePath3)) {
                Assert.False(filePath3.IsFileLocked());

                Assert.Equal(3, document.Paragraphs.Count);
                document.AddParagraph("More paragraphs!");
                Assert.Equal(4, document.Paragraphs.Count);
                document.Save(filePath4);
            }

            Assert.False(filePath1.IsFileLocked());
            Assert.False(filePath2.IsFileLocked());
            Assert.False(filePath3.IsFileLocked());
            Assert.False(filePath4.IsFileLocked());

            using (var document = WordDocument.Load(filePath3)) {
                Assert.Equal(3, document.Paragraphs.Count);
                Assert.False(filePath3.IsFileLocked());
            }
            using (var document = WordDocument.Load(filePath4)) {
                Assert.Equal(4, document.Paragraphs.Count);
                Assert.False(filePath4.IsFileLocked());
            }

            Assert.False(filePath1.IsFileLocked());
            Assert.False(filePath2.IsFileLocked());
            Assert.False(filePath3.IsFileLocked());
            Assert.False(filePath4.IsFileLocked());
        }

        [Fact]
        public void Test_Dispose() {
            var filePath1 = Path.Combine(_directoryWithFiles, "DisposeTesting.docx");
            File.Delete(filePath1);

            Assert.False(File.Exists(filePath1));

            using var document = WordDocument.Create(filePath1);
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";

            document.AddParagraph("This is my test");

            Assert.False(filePath1.IsFileLocked());

            document.Save();

            Assert.False(filePath1.IsFileLocked());
            Assert.True(File.Exists(filePath1));
        }

        [Fact]
        public void Test_SaveToStream() {
            using var document = WordDocument.Create();
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";
            document.AddParagraph("Hello world!");

            using var outputStream = new MemoryStream();
            document.Save(outputStream);

            using (var openXmlDoc = WordprocessingDocument.Open(outputStream, false)) {
                Assert.NotNull(openXmlDoc.MainDocumentPart);
            }
            outputStream.Seek(0, SeekOrigin.Begin);

            using var resultDoc = WordDocument.Load(outputStream);

            Assert.True(resultDoc.BuiltinDocumentProperties.Title == "This is my title");
            Assert.True(resultDoc.BuiltinDocumentProperties.Creator == "Przemysław Kłys");
            Assert.True(resultDoc.BuiltinDocumentProperties.Keywords == "word, docx, test");

            var paragraph = Assert.Single(resultDoc.Paragraphs);
            Assert.Equal("Hello world!", paragraph.Text);
        }


        [Fact]
        public void Test_SaveToStreamAndFile() {
            var filePath = Path.Combine(_directoryWithFiles, "DisposeTesting1.docx");
            File.Delete(filePath);

            Assert.False(File.Exists(filePath));

            using var document = WordDocument.Create();
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";
            document.AddParagraph("Hello world!");

            using var outputStream = new MemoryStream();
            document.Save(outputStream);

            using (var openXmlDoc = WordprocessingDocument.Open(outputStream, false)) {
                Assert.NotNull(openXmlDoc.MainDocumentPart);
            }
            outputStream.Seek(0, SeekOrigin.Begin);

            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                outputStream.CopyTo(fileStream);
            }

            using (var resultDoc = WordDocument.Load(filePath)) {
                Assert.True(resultDoc.BuiltinDocumentProperties.Title == "This is my title");
                Assert.True(resultDoc.BuiltinDocumentProperties.Creator == "Przemysław Kłys");
                Assert.True(resultDoc.BuiltinDocumentProperties.Keywords == "word, docx, test");

                var paragraph = Assert.Single(resultDoc.Paragraphs);
                Assert.Equal("Hello world!", paragraph.Text);

                resultDoc.Save();
            }

            using (var resultDoc = WordDocument.Load(filePath)) {
                Assert.True(resultDoc.BuiltinDocumentProperties.Title == "This is my title");
                Assert.True(resultDoc.BuiltinDocumentProperties.Creator == "Przemysław Kłys");
                Assert.True(resultDoc.BuiltinDocumentProperties.Keywords == "word, docx, test");
            }
        }

        [Fact]
        public void Test_SaveToStreamValidity() {
            using var document = WordDocument.Create();
            document.AddParagraph("Test");

            using var outputStream = new MemoryStream();
            document.Save(outputStream);

            Assert.Equal(0, outputStream.Position);

            using var openXmlDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(outputStream, false);
            Assert.NotNull(openXmlDoc.MainDocumentPart);
        }

        [Fact]
        public void Test_CreateDocumentInStream() {
            using var stream = new MemoryStream();
            using var document = WordDocument.Create(stream);
            document.AddParagraph("Stream paragraph");

            document.Save(stream);

            using (var openXmlDoc = WordprocessingDocument.Open(stream, false)) {
                Assert.NotNull(openXmlDoc.MainDocumentPart);
            }
            stream.Seek(0, SeekOrigin.Begin);

            using var loadedDoc = WordDocument.Load(stream);
            var paragraph = Assert.Single(loadedDoc.Paragraphs);
            Assert.Equal("Stream paragraph", paragraph.Text);
        }

        [Fact]
        public void Test_SaveWithoutFilePath_UsesOriginalStream() {
            using var stream = new MemoryStream();
            using (var document = WordDocument.Create(stream)) {
                document.AddParagraph("Hello original stream");
                document.Save();
            }

            using (var openXmlDoc = WordprocessingDocument.Open(stream, false)) {
                Assert.NotNull(openXmlDoc.MainDocumentPart);
            }
            stream.Seek(0, SeekOrigin.Begin);

            using var loadedDoc = WordDocument.Load(stream);
            var paragraph = Assert.Single(loadedDoc.Paragraphs);
            Assert.Equal("Hello original stream", paragraph.Text);
        }
       
        [Fact]
        public void Test_SaveReadOnlyDocument_ThrowsInvalidOperationException() {
            var filePath = Path.Combine(_directoryWithFiles, "ReadOnlyDocument.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test");
                document.Save();
            }

            using var readOnlyDocument = WordDocument.Load(filePath, readOnly: true);
            Assert.Throws<InvalidOperationException>(() => readOnlyDocument.Save());
            using var outputStream = new MemoryStream();
            Assert.Throws<InvalidOperationException>(() => readOnlyDocument.Save(outputStream));
        }

    }

}
