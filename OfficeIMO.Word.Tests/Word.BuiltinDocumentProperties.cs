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

                document.Properties.Creator = "Przemysław Kłys";
                document.Properties.Title = "Test document";

                document.Properties.Description = "Test document used during testing";
                document.Properties.Category = "Some Category";
                document.Properties.Subject = "Some Subject";
                document.Properties.LastModifiedBy = "Zenon Jaskuła";
                document.Properties.Keywords = "Keyword1,Keyword2";

                document.Properties.Created = new DateTime(2000, 05, 01);
                document.Properties.Modified = new DateTime(2000, 06, 01);
                document.Properties.LastPrinted = new DateTime(2000, 07, 01);
                document.Properties.Version = "0.1.0";
                document.Properties.Revision = "0.5.0";

                Assert.True(document.Properties.Creator == "Przemysław Kłys", "Wrong creator");
                Assert.True(document.Properties.Title == "Test document", "Wrong Title");
                Assert.True(document.Properties.Description == "Test document used during testing", "Wrong Description");
                Assert.True(document.Properties.Category == "Some Category", "Wrong Category");
                Assert.True(document.Properties.Subject == "Some Subject", "Wrong Subject");
                Assert.True(document.Properties.LastModifiedBy == "Zenon Jaskuła", "Wrong LastModifiedBy");
                Assert.True(document.Properties.Keywords == "Keyword1,Keyword2", "Wrong Keywords");
                Assert.True(document.Properties.Created == new DateTime(2000, 05, 01), "Wrong Created");
                Assert.True(document.Properties.Modified == new DateTime(2000, 06, 01), "Wrong Modified");
                Assert.True(document.Properties.LastPrinted == new DateTime(2000, 07, 01), "Wrong LastPrinted");
                Assert.True(document.Properties.Version == "0.1.0", "Wrong Version");
                Assert.True(document.Properties.Revision == "0.5.0", "Wrong Revision");
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentCreationWithProperties.docx"))) {

                Assert.True(document.Properties.Creator == "Przemysław Kłys", "Wrong creator");
                Assert.True(document.Properties.Title == "Test document", "Wrong Title");
                Assert.True(document.Properties.Description == "Test document used during testing", "Wrong Description");
                Assert.True(document.Properties.Category == "Some Category", "Wrong Category");
                Assert.True(document.Properties.Subject == "Some Subject", "Wrong Subject");
                Assert.True(document.Properties.LastModifiedBy == "Zenon Jaskuła", "Wrong LastModifiedBy");
                Assert.True(document.Properties.Keywords == "Keyword1,Keyword2", "Wrong Keywords");
                Assert.True(document.Properties.Created == new DateTime(2000, 05, 01), "Wrong Created");
                Assert.True(document.Properties.Modified == new DateTime(2000, 06, 01), "Wrong Modified");
                Assert.True(document.Properties.LastPrinted == new DateTime(2000, 07, 01), "Wrong LastPrinted");
                Assert.True(document.Properties.Version == "0.1.0", "Wrong Version");
                Assert.True(document.Properties.Revision == "0.5.0", "Wrong Revision");
                document.Properties.Modified = new DateTime(2001, 06, 01);
                document.Properties.Creator = "Evotec Services";
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentCreationWithProperties.docx"))) {

                Assert.True(document.Properties.Creator == "Evotec Services", "Wrong creator");
                Assert.True(document.Properties.Title == "Test document", "Wrong Title");
                Assert.True(document.Properties.Description == "Test document used during testing", "Wrong Description");
                Assert.True(document.Properties.Category == "Some Category", "Wrong Category");
                Assert.True(document.Properties.Subject == "Some Subject", "Wrong Subject");
                Assert.True(document.Properties.LastModifiedBy == "Zenon Jaskuła", "Wrong LastModifiedBy");
                Assert.True(document.Properties.Keywords == "Keyword1,Keyword2", "Wrong Keywords");
                Assert.True(document.Properties.Created == new DateTime(2000, 05, 01), "Wrong Created");
                Assert.True(document.Properties.Modified == new DateTime(2001, 06, 01), "Wrong Modified");
                Assert.True(document.Properties.LastPrinted == new DateTime(2000, 07, 01), "Wrong LastPrinted");
                Assert.True(document.Properties.Version == "0.1.0", "Wrong Version");
                Assert.True(document.Properties.Revision == "0.5.0", "Wrong Revision");
                document.Save();
            }
        }
    }
}
