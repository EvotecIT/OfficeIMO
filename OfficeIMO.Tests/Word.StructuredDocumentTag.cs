using OfficeIMO.Word;
using Xunit;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingStructuredDocumentTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello world", "Alias1");

                Assert.True(document.StructuredDocumentTags.Count == 1);
                Assert.True(document.ParagraphsStructuredDocumentTags.Count == 1);
                Assert.Equal("Hello world", sdt.Text);

                document.Save(false);

                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.StructuredDocumentTags.Count == 1);
                Assert.Equal("Hello world", document.StructuredDocumentTags[0].Text);

                document.StructuredDocumentTags[0].Text = "Changed";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Changed", document.StructuredDocumentTags[0].Text);
            }
        }

        [Fact]
        public void Test_StructuredDocumentTagWithTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlTag.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello", "Alias1", "Tag1");

                Assert.Equal("Tag1", sdt.Tag);
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var loaded = document.GetStructuredDocumentTagByTag("Tag1");
                Assert.NotNull(loaded);
                Assert.Equal("Hello", loaded.Text);

                loaded.Text = "Updated";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Updated", document.StructuredDocumentTags[0].Text);
                Assert.Equal("Tag1", document.StructuredDocumentTags[0].Tag);
            }
        }

        [Fact]
        public void Test_StructuredDocumentTagGetByAlias() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlAlias.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello", "Alias100", "Tag100");

                Assert.NotNull(document.GetStructuredDocumentTagByAlias("Alias100"));
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var alias = document.GetStructuredDocumentTagByAlias("Alias100");
                Assert.NotNull(alias);
                alias.Text = "Updated";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Updated", document.StructuredDocumentTags[0].Text);
                Assert.Equal("Tag100", document.StructuredDocumentTags[0].Tag);
            }
        }

        [Fact]
        public void Test_StructuredDocumentTagFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var para1 = document.AddParagraph("Formatted:");
                var sdt1 = para1.AddStructuredDocumentTag("Styled", "AliasProps", "TagProps");
                sdt1.Bold = true;
                sdt1.Italic = true;
                sdt1.Underline = UnderlineValues.Single;
                sdt1.FontFamily = "Calibri";
                sdt1.FontSize = 12;
                sdt1.ColorHex = "2F5597";
                sdt1.Highlight = HighlightColorValues.Yellow;

                var para2 = document.AddParagraph("Formatted fluent:");
                para2.AddStructuredDocumentTag("Fluent", "AliasFluent", "TagFluent")
                    .SetBold()
                    .SetItalic()
                    .SetUnderline(UnderlineValues.Single)
                    .SetFontFamily("Calibri")
                    .SetFontSize(14)
                    .SetColorHex("C00000")
                    .SetHighlight(HighlightColorValues.LightGray);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var props = document.GetStructuredDocumentTagByTag("TagProps");
                Assert.NotNull(props);
                Assert.True(props!.Bold);
                Assert.True(props.Italic);
                Assert.Equal(UnderlineValues.Single, props.Underline);
                Assert.Equal(12, props.FontSize);
                Assert.Equal("Calibri", props.FontFamily);
                Assert.Equal("2f5597", props.ColorHex);
                Assert.Equal(HighlightColorValues.Yellow, props.Highlight);

                var fluent = document.GetStructuredDocumentTagByTag("TagFluent");
                Assert.NotNull(fluent);
                Assert.True(fluent!.Bold);
                Assert.True(fluent.Italic);
                Assert.Equal(UnderlineValues.Single, fluent.Underline);
                Assert.Equal(14, fluent.FontSize);
                Assert.Equal("Calibri", fluent.FontFamily);
                Assert.Equal("c00000", fluent.ColorHex);
                Assert.Equal(HighlightColorValues.LightGray, fluent.Highlight);
            }
        }

        [Fact]
        public void Test_SettingTextOnEmptyStructuredDocumentTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithEmptySdt.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdtRun = new SdtRun();

                var properties = new SdtProperties();
                properties.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(new System.Random().Next(1, int.MaxValue)) });
                sdtRun.Append(properties);
                sdtRun.Append(new SdtContentRun());

                var paragraph = new Paragraph(sdtRun);
                var wordParagraph = new WordParagraph(document, paragraph, sdtRun);
                document.AddParagraph(wordParagraph);

                var sdt = wordParagraph.StructuredDocumentTag;
                Assert.NotNull(sdt);
                Assert.Null(sdt!.Text);

                sdt.Text = "New text";

                Assert.Equal("New text", sdt.Text);
                var run = sdtRun.SdtContentRun?.GetFirstChild<Run>();
                Assert.NotNull(run);
                Assert.Equal("New text", run!.GetFirstChild<Text>()?.Text);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.StructuredDocumentTags);
                Assert.Equal("New text", document.StructuredDocumentTags[0].Text);
            }
        }
    }
}
