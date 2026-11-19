using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TablesOfContentsOnlyKeepsAppendedFieldResults() {
            using WordDocument destination = WordDocument.Create();
            destination.Properties.Title = "Brief A";
            destination.CustomProperties.Add("Classification", new WordCustomProperty("OFFICIAL"));
            destination.AddTableOfContent();
            destination.AddParagraph("Brief A heading").Style = WordParagraphStyles.Heading1;

            using WordDocument source = WordDocument.Create();
            source.Properties.Title = "Brief B";
            source.CustomProperties.Add("Classification", new WordCustomProperty("PROTECTED"));
            source.AddParagraph("Brief B heading").Style = WordParagraphStyles.Heading1;
            Paragraph fields = source.AddParagraph("Fields: ")._paragraph;
            fields.Append(new SimpleField(new Run(new Text("PROTECTED"))) { Instruction = " DOCPROPERTY Classification " });
            fields.Append(new SimpleField(new Run(new Text("Brief B"))) { Instruction = " TITLE " });
            fields.Append(new SimpleField(new Run(new Text("1 January 2020"))) { Instruction = " DATE " });

            destination.AppendDocument(source);
            WordFieldUpdateReport report = destination.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                TablesOfContentsOnly = true,
                CurrentDateTime = new DateTime(2030, 6, 7)
            });

            Dictionary<string, int> pages = ReadTocEntryPages(destination.TableOfContent!.SdtBlock);
            Assert.Equal(new[] { "Brief A heading", "Brief B heading" }, pages.Keys.ToArray());
            Assert.Equal(WordFieldUpdateStatus.Updated, report.Results.Single(_ => _.FieldType == WordFieldType.TOC).Status);
            var pageReferences = report.Results.Where(_ => _.FieldType == WordFieldType.PageRef).ToList();
            Assert.Equal(2, pageReferences.Count);
            Assert.All(pageReferences, _ => Assert.Equal(WordFieldUpdateStatus.Updated, _.Status));

            // The appended fields keep the results their source gave them rather than the destination's values.
            Dictionary<string, string> appended = destination._document.Body!
                .Descendants<SimpleField>()
                .ToDictionary(_ => _.Instruction!.Value!.Trim(), _ => _.InnerText);
            Assert.Equal("PROTECTED", appended["DOCPROPERTY Classification"]);
            Assert.Equal("Brief B", appended["TITLE"]);
            Assert.Equal("1 January 2020", appended["DATE"]);
            Assert.All(
                report.Results.Where(_ => _.FieldType is WordFieldType.DocProperty or WordFieldType.Title or WordFieldType.Date),
                _ => Assert.Equal(WordFieldUpdateStatus.Skipped, _.Status));

            // The entries are final, so the table no longer asks Word to rebuild it on open.
            Assert.False(destination.Settings.UpdateFieldsOnOpen);
        }

        [Fact]
        public void Test_TablesOfContentsOnlyCannotBeCombinedWithDateTimeFieldsOnly() {
            using WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddParagraph("Chapter").Style = WordParagraphStyles.Heading1;
            string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

            Assert.Throws<InvalidOperationException>(() => document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                DateTimeFieldsOnly = true,
                TablesOfContentsOnly = true
            }));

            Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        }
    }
}
