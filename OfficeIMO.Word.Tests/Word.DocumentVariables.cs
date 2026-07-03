using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DocumentVariables() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithVariables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.SetDocumentVariable("TestVar", "Value1");
                document.SetDocumentVariable("AnotherVar", "123");
                Assert.True(document.GetDocumentVariable("TestVar") == "Value1");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.GetDocumentVariable("AnotherVar") == "123");
                document.SetDocumentVariable("TestVar", "Updated");
                document.RemoveDocumentVariable("AnotherVar");
                Assert.True(document.HasDocumentVariables);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.GetDocumentVariable("TestVar") == "Updated");
                Assert.False(document.DocumentVariables.ContainsKey("AnotherVar"));
                document.RemoveDocumentVariableAt(0);
                Assert.False(document.HasDocumentVariables);
                document.Save();
            }
        }
    }
}
