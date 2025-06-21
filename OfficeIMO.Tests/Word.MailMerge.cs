using System.Collections.Generic;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MailMerge_ReplacesFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMerge.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" })
                    .AddText("!");

                var values = new Dictionary<string, string> { { "Name", "Alice" } };
                WordMailMerge.Execute(document, values);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                string xml = document._document.MainDocumentPart.Document.InnerText;
                Assert.Contains("Alice", xml);
                Assert.DoesNotContain("MERGEFIELD", xml);
            }
        }

        [Fact]
        public void Test_MailMerge_KeepFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeKeep.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" })
                    .AddText("!");

                var values = new Dictionary<string, string> { { "Name", "Bob" } };
                WordMailMerge.Execute(document, values, removeFields: false);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.Fields);
                Assert.Equal("Bob", document.Fields[0].Text);
            }
        }
    }
}
