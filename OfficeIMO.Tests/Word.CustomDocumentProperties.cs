using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_SimpleWordDocumentCreationWithCustomProperties() {
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

                document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = new DateTime(2000,1,1, 5, 5, 5) });
                document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
                document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));
                document.CustomDocumentProperties.Add("Number", new WordCustomProperty(1500));
                document.CustomDocumentProperties.Add("NumberDouble", new WordCustomProperty(15.00));

                document.CustomDocumentProperties.Add("TestDifferentWay", new WordCustomProperty("String", PropertyTypes.Text));
                document.CustomDocumentProperties.Add("TestDifferentWayNumber", new WordCustomProperty(15, PropertyTypes.NumberInteger));

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

                // check custom properties
                Assert.True((DateTime) document.CustomDocumentProperties["TestProperty"].Value == new DateTime(2000, 1, 1, 5, 5, 5));
                Assert.True(document.CustomDocumentProperties["TestProperty"].Date == new DateTime(2000, 1, 1, 5, 5, 5));
                Assert.True(((string)document.CustomDocumentProperties["MyName"].Value) == "Some text");
                Assert.True(document.CustomDocumentProperties["MyName"].Text == "Some text");
                Assert.True((bool) document.CustomDocumentProperties["IsTodayGreatDay"].Value == true);
                Assert.True(document.CustomDocumentProperties["IsTodayGreatDay"].Bool == true);

                Assert.True(document.CustomDocumentProperties["Number"].NumberInteger == 1500);
                Assert.True((int) document.CustomDocumentProperties["Number"].Value == 1500);

                Assert.True(document.CustomDocumentProperties["NumberDouble"].NumberDouble == 15.00);
                Assert.True((double)document.CustomDocumentProperties["NumberDouble"].Value == 15.00);
                
                Assert.True(((string)document.CustomDocumentProperties["TestDifferentWay"].Value) == "String");
                Assert.True(document.CustomDocumentProperties["TestDifferentWay"].Text == "String");
                Assert.True(document.CustomDocumentProperties["TestDifferentWayNumber"].NumberInteger == 15);
                Assert.True((int)document.CustomDocumentProperties["TestDifferentWayNumber"].Value == 15);

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

                // check custom properties
                Assert.True((DateTime)document.CustomDocumentProperties["TestProperty"].Value == new DateTime(2000, 1, 1, 5, 5, 5));
                Assert.True(document.CustomDocumentProperties["TestProperty"].Date == new DateTime(2000, 1, 1,5,5,5));
                Assert.True(((string) document.CustomDocumentProperties["MyName"].Value) == "Some text");
                Assert.True(document.CustomDocumentProperties["MyName"].Text == "Some text");
                Assert.True((bool)document.CustomDocumentProperties["IsTodayGreatDay"].Value == true);
                Assert.True(document.CustomDocumentProperties["IsTodayGreatDay"].Bool == true);

                Assert.True(document.CustomDocumentProperties["Number"].NumberInteger == 1500);
                Assert.True((int)document.CustomDocumentProperties["Number"].Value == 1500);

                Assert.True(document.CustomDocumentProperties["NumberDouble"].NumberDouble == 15.00);
                Assert.True((double)document.CustomDocumentProperties["NumberDouble"].Value == 15.00);

                Assert.True(((string)document.CustomDocumentProperties["TestDifferentWay"].Value) == "String");
                Assert.True(document.CustomDocumentProperties["TestDifferentWay"].Text == "String");
                Assert.True(document.CustomDocumentProperties["TestDifferentWayNumber"].NumberInteger == 15);
                Assert.True((int)document.CustomDocumentProperties["TestDifferentWayNumber"].Value == 15);

                document.CustomDocumentProperties["NumberDouble"].Value = 6.15;
                document.CustomDocumentProperties["TestProperty"].Value = new DateTime(2010, 1, 1, 5, 5, 5);

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


                // check custom properties
                Assert.True((DateTime)document.CustomDocumentProperties["TestProperty"].Value == new DateTime(2010, 1, 1, 5, 5, 5));
                Assert.True(document.CustomDocumentProperties["TestProperty"].Date == new DateTime(2010, 1, 1, 5, 5, 5));
                Assert.True(((string)document.CustomDocumentProperties["MyName"].Value) == "Some text");
                Assert.True(document.CustomDocumentProperties["MyName"].Text == "Some text");
                Assert.True((bool)document.CustomDocumentProperties["IsTodayGreatDay"].Value == true);
                Assert.True(document.CustomDocumentProperties["IsTodayGreatDay"].Bool == true);

                Assert.True(document.CustomDocumentProperties["Number"].NumberInteger == 1500);
                Assert.True((int)document.CustomDocumentProperties["Number"].Value == 1500);

                Assert.True(document.CustomDocumentProperties["NumberDouble"].NumberDouble == 6.15);
                Assert.True((double)document.CustomDocumentProperties["NumberDouble"].Value == 6.15);

                Assert.True(((string)document.CustomDocumentProperties["TestDifferentWay"].Value) == "String");
                Assert.True(document.CustomDocumentProperties["TestDifferentWay"].Text == "String");
                Assert.True(document.CustomDocumentProperties["TestDifferentWayNumber"].NumberInteger == 15);
                Assert.True((int)document.CustomDocumentProperties["TestDifferentWayNumber"].Value == 15);

                document.Save(false);
            }
        }
    }
}
