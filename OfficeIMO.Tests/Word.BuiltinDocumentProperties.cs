using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_SimpleWordDocumentCreationWithProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "SimpleWordDocumentCreationWithProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Title = "Test document";

                document.BuiltinDocumentProperties.Description = "Test document used during testing";
                document.BuiltinDocumentProperties.Category = "Some Category";
                document.BuiltinDocumentProperties.Subject = "Some Subject";
                document.BuiltinDocumentProperties.LastModifiedBy = "Zenon Jaskuła";
                document.BuiltinDocumentProperties.Keywords = "Keyword1,Keyword2";

                document.BuiltinDocumentProperties.Created = new DateTime(2000, 05, 01);
                document.BuiltinDocumentProperties.Modified = new DateTime(2000, 06, 01);
                document.BuiltinDocumentProperties.LastPrinted = new DateTime(2000, 07, 01);
                document.BuiltinDocumentProperties.Version = "0.1.0";
                document.BuiltinDocumentProperties.Revision = "0.5.0";

                Assert.True(document.BuiltinDocumentProperties.Creator == "Przemysław Kłys", "Wrong creator");
                Assert.True(document.BuiltinDocumentProperties.Title == "Test document", "Wrong Title");
                Assert.True(document.BuiltinDocumentProperties.Description == "Test document used during testing", "Wrong Description");
                Assert.True(document.BuiltinDocumentProperties.Category == "Some Category", "Wrong Category");
                Assert.True(document.BuiltinDocumentProperties.Subject == "Some Subject", "Wrong Subject");
                Assert.True(document.BuiltinDocumentProperties.LastModifiedBy == "Zenon Jaskuła", "Wrong LastModifiedBy");
                Assert.True(document.BuiltinDocumentProperties.Keywords == "Keyword1,Keyword2", "Wrong Keywords");
                Assert.True(document.BuiltinDocumentProperties.Created == new DateTime(2000, 05, 01), "Wrong Created");
                Assert.True(document.BuiltinDocumentProperties.Modified == new DateTime(2000, 06, 01), "Wrong Modified");
                Assert.True(document.BuiltinDocumentProperties.LastPrinted == new DateTime(2000, 07, 01), "Wrong LastPrinted");
                Assert.True(document.BuiltinDocumentProperties.Version == "0.1.0", "Wrong Version");
                Assert.True(document.BuiltinDocumentProperties.Revision == "0.5.0", "Wrong Revision");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentCreationWithProperties.docx"))) {

                Assert.True(document.BuiltinDocumentProperties.Creator == "Przemysław Kłys", "Wrong creator");
                Assert.True(document.BuiltinDocumentProperties.Title == "Test document", "Wrong Title");
                Assert.True(document.BuiltinDocumentProperties.Description == "Test document used during testing", "Wrong Description");
                Assert.True(document.BuiltinDocumentProperties.Category == "Some Category", "Wrong Category");
                Assert.True(document.BuiltinDocumentProperties.Subject == "Some Subject", "Wrong Subject");
                Assert.True(document.BuiltinDocumentProperties.LastModifiedBy == "Zenon Jaskuła", "Wrong LastModifiedBy");
                Assert.True(document.BuiltinDocumentProperties.Keywords == "Keyword1,Keyword2", "Wrong Keywords");
                Assert.True(document.BuiltinDocumentProperties.Created == new DateTime(2000, 05, 01), "Wrong Created");
                Assert.True(document.BuiltinDocumentProperties.Modified == new DateTime(2000, 06, 01), "Wrong Modified");
                Assert.True(document.BuiltinDocumentProperties.LastPrinted == new DateTime(2000, 07, 01), "Wrong LastPrinted");
                Assert.True(document.BuiltinDocumentProperties.Version == "0.1.0", "Wrong Version");
                Assert.True(document.BuiltinDocumentProperties.Revision == "0.5.0", "Wrong Revision");
                document.BuiltinDocumentProperties.Modified = new DateTime(2001, 06, 01);
                document.BuiltinDocumentProperties.Creator = "Evotec Services";
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentCreationWithProperties.docx"))) {

                Assert.True(document.BuiltinDocumentProperties.Creator == "Evotec Services", "Wrong creator");
                Assert.True(document.BuiltinDocumentProperties.Title == "Test document", "Wrong Title");
                Assert.True(document.BuiltinDocumentProperties.Description == "Test document used during testing", "Wrong Description");
                Assert.True(document.BuiltinDocumentProperties.Category == "Some Category", "Wrong Category");
                Assert.True(document.BuiltinDocumentProperties.Subject == "Some Subject", "Wrong Subject");
                Assert.True(document.BuiltinDocumentProperties.LastModifiedBy == "Zenon Jaskuła", "Wrong LastModifiedBy");
                Assert.True(document.BuiltinDocumentProperties.Keywords == "Keyword1,Keyword2", "Wrong Keywords");
                Assert.True(document.BuiltinDocumentProperties.Created == new DateTime(2000, 05, 01), "Wrong Created");
                Assert.True(document.BuiltinDocumentProperties.Modified == new DateTime(2001, 06, 01), "Wrong Modified");
                Assert.True(document.BuiltinDocumentProperties.LastPrinted == new DateTime(2000, 07, 01), "Wrong LastPrinted");
                Assert.True(document.BuiltinDocumentProperties.Version == "0.1.0", "Wrong Version");
                Assert.True(document.BuiltinDocumentProperties.Revision == "0.5.0", "Wrong Revision");
                document.Save(false);
            }
        }
    }
}
