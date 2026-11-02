using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests related to custom document properties.
    /// </summary>
    public partial class Word {
        /// <summary>
        /// Creates a document with custom properties and verifies values.
        /// </summary>
        [Fact]
        public void Test_SimpleWordDocumentCreationWithCustomProperties() {
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

                document.CustomProperties.Add("TestProperty", new WordCustomProperty { Value = new DateTime(2000,1,1, 5, 5, 5) });
                document.CustomProperties.Add("MyName", new WordCustomProperty("Some text"));
                document.CustomProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));
                document.CustomProperties.Add("Number", new WordCustomProperty(1500));
                document.CustomProperties.Add("NumberDouble", new WordCustomProperty(15.00));

                document.CustomProperties.Add("TestDifferentWay", new WordCustomProperty("String", WordCustomPropertyType.Text));
                document.CustomProperties.Add("TestDifferentWayNumber", new WordCustomProperty(15, WordCustomPropertyType.NumberInteger));

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

                // check custom properties
                Assert.True((DateTime) document.CustomProperties["TestProperty"].Value == new DateTime(2000, 1, 1, 5, 5, 5));
                Assert.True(document.CustomProperties["TestProperty"].Date == new DateTime(2000, 1, 1, 5, 5, 5));
                Assert.True(((string)document.CustomProperties["MyName"].Value) == "Some text");
                Assert.True(document.CustomProperties["MyName"].Text == "Some text");
                Assert.True((bool) document.CustomProperties["IsTodayGreatDay"].Value == true);
                Assert.True(document.CustomProperties["IsTodayGreatDay"].Bool == true);

                Assert.True(document.CustomProperties["Number"].NumberInteger == 1500);
                Assert.True((int) document.CustomProperties["Number"].Value == 1500);

                Assert.True(document.CustomProperties["NumberDouble"].NumberDouble == 15.00);
                Assert.True((double)document.CustomProperties["NumberDouble"].Value == 15.00);
                
                Assert.True(((string)document.CustomProperties["TestDifferentWay"].Value) == "String");
                Assert.True(document.CustomProperties["TestDifferentWay"].Text == "String");
                Assert.True(document.CustomProperties["TestDifferentWayNumber"].NumberInteger == 15);
                Assert.True((int)document.CustomProperties["TestDifferentWayNumber"].Value == 15);

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

                // check custom properties
                Assert.True((DateTime)document.CustomProperties["TestProperty"].Value == new DateTime(2000, 1, 1, 5, 5, 5));
                Assert.True(document.CustomProperties["TestProperty"].Date == new DateTime(2000, 1, 1,5,5,5));
                Assert.True(((string) document.CustomProperties["MyName"].Value) == "Some text");
                Assert.True(document.CustomProperties["MyName"].Text == "Some text");
                Assert.True((bool)document.CustomProperties["IsTodayGreatDay"].Value == true);
                Assert.True(document.CustomProperties["IsTodayGreatDay"].Bool == true);

                Assert.True(document.CustomProperties["Number"].NumberInteger == 1500);
                Assert.True((int)document.CustomProperties["Number"].Value == 1500);

                Assert.True(document.CustomProperties["NumberDouble"].NumberDouble == 15.00);
                Assert.True((double)document.CustomProperties["NumberDouble"].Value == 15.00);

                Assert.True(((string)document.CustomProperties["TestDifferentWay"].Value) == "String");
                Assert.True(document.CustomProperties["TestDifferentWay"].Text == "String");
                Assert.True(document.CustomProperties["TestDifferentWayNumber"].NumberInteger == 15);
                Assert.True((int)document.CustomProperties["TestDifferentWayNumber"].Value == 15);

                document.CustomProperties["NumberDouble"].Value = 6.15;
                document.CustomProperties["TestProperty"].Value = new DateTime(2010, 1, 1, 5, 5, 5);

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


                // check custom properties
                Assert.True((DateTime)document.CustomProperties["TestProperty"].Value == new DateTime(2010, 1, 1, 5, 5, 5));
                Assert.True(document.CustomProperties["TestProperty"].Date == new DateTime(2010, 1, 1, 5, 5, 5));
                Assert.True(((string)document.CustomProperties["MyName"].Value) == "Some text");
                Assert.True(document.CustomProperties["MyName"].Text == "Some text");
                Assert.True((bool)document.CustomProperties["IsTodayGreatDay"].Value == true);
                Assert.True(document.CustomProperties["IsTodayGreatDay"].Bool == true);

                Assert.True(document.CustomProperties["Number"].NumberInteger == 1500);
                Assert.True((int)document.CustomProperties["Number"].Value == 1500);

                Assert.True(document.CustomProperties["NumberDouble"].NumberDouble == 6.15);
                Assert.True((double)document.CustomProperties["NumberDouble"].Value == 6.15);

                Assert.True(((string)document.CustomProperties["TestDifferentWay"].Value) == "String");
                Assert.True(document.CustomProperties["TestDifferentWay"].Text == "String");
                Assert.True(document.CustomProperties["TestDifferentWayNumber"].NumberInteger == 15);
                Assert.True((int)document.CustomProperties["TestDifferentWayNumber"].Value == 15);

                document.Save();
            }
        }

        [Fact]
        public void Test_WordCustomDocumentProperties_RoundTripBinaryValue() {
            string filePath = Path.Combine(_directoryWithFiles, "WordCustomDocumentProperties.Binary.docx");
            byte[] payload = { 0x00, 0x01, 0x42, 0x80, 0xff };

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Binary custom property");
                document.CustomProperties.Add("BinaryPayload", new WordCustomProperty(payload));
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                CustomDocumentProperty binary = package.CustomFilePropertiesPart!.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "BinaryPayload");
                Assert.NotNull(binary.VTBlob);
                Assert.Equal(Convert.ToBase64String(payload), binary.VTBlob!.Text);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordCustomProperty property = document.CustomProperties["BinaryPayload"];
                Assert.Equal(WordCustomPropertyType.Binary, property.PropertyType);
                Assert.Equal(payload, property.Binary);
            }
        }
    }
}
