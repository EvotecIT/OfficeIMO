using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
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
                Assert.True(filePath1.IsFileLocked());
            }
            using (var document = WordDocument.Load(filePath2)) {
                Assert.Equal(2, document.Paragraphs.Count);
                Assert.True(filePath2.IsFileLocked());
            }
            using (var document = WordDocument.Load(filePath3)) {
                Assert.True(filePath3.IsFileLocked());

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
                Assert.True(filePath3.IsFileLocked());
            }
            using (var document = WordDocument.Load(filePath4)) {
                Assert.Equal(4, document.Paragraphs.Count);
                Assert.True(filePath4.IsFileLocked());
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

            var document = WordDocument.Create(filePath1);
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";

            document.AddParagraph("This is my test");

            Assert.True(filePath1.IsFileLocked());

            document.Save();

            Assert.False(filePath1.IsFileLocked());

            document.Dispose();

            Assert.False(filePath1.IsFileLocked());
            Assert.True(File.Exists(filePath1));
        }

        [Fact]
        public void Test_SaveToStream() {
            var document = WordDocument.Create();
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";
            document.AddParagraph("Hello world!");

            using var outputStream = new MemoryStream();
            document.Save(outputStream);

            var resultDoc = WordDocument.Load(outputStream);

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

            var document = WordDocument.Create();
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";
            document.AddParagraph("Hello world!");

            using var outputStream = new MemoryStream();
            document.Save(outputStream);

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

    }

}
