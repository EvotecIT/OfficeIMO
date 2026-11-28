using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class WordFieldInventoryTests {
        private readonly string _directoryWithFiles;

        public WordFieldInventoryTests() {
            _directoryWithFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempDocuments2", Guid.NewGuid().ToString("N"));
            Word.Setup(_directoryWithFiles);
        }

        [Fact]
        public void Test_InspectFields_ReadsSimpleComplexSplitNestedAndUnsupportedFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldInventory.Basic.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Author: ").AddField(WordFieldType.Author);
                AddSplitComplexField(
                    document.AddParagraph()._paragraph,
                    " REF ",
                    " \"Bookmark1\" \\h ",
                    "Referenced Heading",
                    dirty: true,
                    locked: true);
                AddNestedComplexFields(document.AddParagraph()._paragraph);
                document.AddParagraph()._paragraph.Append(BuildSimpleField(" SILLYFIELD value ", "Unknown"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Equal(5, fields.Count);

                WordFieldInfo simpleAuthor = fields.Single(field =>
                    field.FieldType == WordFieldType.Author &&
                    field.Representation == WordFieldRepresentation.Simple);
                Assert.Equal(0, simpleAuthor.NestingLevel);
                Assert.True(simpleAuthor.IsParsed);

                WordFieldInfo reference = fields.Single(field => field.FieldType == WordFieldType.Ref);
                Assert.Equal(WordFieldRepresentation.Complex, reference.Representation);
                Assert.Equal(" REF  \"Bookmark1\" \\h ", reference.InstructionText);
                Assert.Equal(new[] { "\"Bookmark1\"" }, reference.Instructions);
                Assert.Equal(new[] { "\\h" }, reference.Switches);
                Assert.Equal("Referenced Heading", reference.ResultText);
                Assert.True(reference.IsDirty);
                Assert.True(reference.IsLocked);

                WordFieldInfo quote = fields.Single(field => field.FieldType == WordFieldType.Quote);
                Assert.Equal(0, quote.NestingLevel);
                Assert.Equal("Outer start Nested Author outer end", quote.ResultText);

                WordFieldInfo nestedAuthor = fields.Single(field =>
                    field.FieldType == WordFieldType.Author &&
                    field.Representation == WordFieldRepresentation.Complex);
                Assert.Equal(1, nestedAuthor.NestingLevel);
                Assert.Equal("Nested Author", nestedAuthor.ResultText);

                WordFieldInfo unsupported = fields.Single(field => field.FieldType == null);
                Assert.False(unsupported.IsParsed);
                Assert.Contains(unsupported.UnsupportedParseDetails, detail => detail.Contains("couldn't be processed", StringComparison.OrdinalIgnoreCase));

                WordFeatureFinding fieldFinding = Assert.Single(document.InspectFeatures().FindFeatures("Fields"));
                Assert.Contains(fieldFinding.Details, detail => detail == "Simple fields: 2");
                Assert.Contains(fieldFinding.Details, detail => detail == "Complex fields: 3");
                Assert.Contains(fieldFinding.Details, detail => detail.Contains("Parsed field types:", StringComparison.Ordinal));
                Assert.Contains(fieldFinding.Details, detail => detail == "Field parser diagnostics: 1");
            }
        }

        [Fact]
        public void Test_InspectFields_ReadsSameRunComplexFieldMarkers() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldInventory.SameRunComplexMarkers.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Paragraph paragraph = document.AddParagraph("Value: ")._paragraph;
                paragraph.Append(new Run(
                    new FieldChar { FieldCharType = FieldCharValues.Begin },
                    new FieldCode(" QUOTE \"fresh\" ") { Space = SpaceProcessingModeValues.Preserve },
                    new FieldChar { FieldCharType = FieldCharValues.Separate },
                    new Text("cached") { Space = SpaceProcessingModeValues.Preserve },
                    new FieldChar { FieldCharType = FieldCharValues.End }));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldInfo field = Assert.Single(document.InspectFields(), item => item.FieldType == WordFieldType.Quote);

                Assert.Equal(WordFieldRepresentation.Complex, field.Representation);
                Assert.Equal(" QUOTE \"fresh\" ", field.InstructionText);
                Assert.Equal("cached", field.ResultText);
                Assert.Empty(field.UnsupportedParseDetails);
            }
        }

        [Fact]
        public void Test_InspectFields_ReadsFieldsAcrossPartsAndContainers() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldInventory.PartsAndContainers.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].AddField(WordFieldType.Page);

                _ = document.PageSetup.Header;
                RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                    .AddParagraph("Header title: ")
                    .AddField(WordFieldType.Title);
                RequireSectionFooter(document, 0, HeaderFooterValues.Default)
                    .AddParagraph("Footer pages: ")
                    .AddField(WordFieldType.NumPages);

                document.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote anchor").AddEndNote("Endnote body");

                WordTextBox textBox = document.AddTextBox("Text box host");
                textBox.Paragraphs[0].AddField(WordFieldType.Subject);

                AddContentControlField(document, " KEYWORDS ", "Keyword result");
                document.Save();
            }

            AppendNoteFields(filePath);

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field =>
                    field.LocationKind == WordFieldLocationKind.Body &&
                    field.FieldType == WordFieldType.Page &&
                    field.IsInTable);
                Assert.Contains(fields, field =>
                    field.LocationKind == WordFieldLocationKind.Header &&
                    field.FieldType == WordFieldType.Title &&
                    field.PartUri.Contains("header", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(fields, field =>
                    field.LocationKind == WordFieldLocationKind.Footer &&
                    field.FieldType == WordFieldType.NumPages &&
                    field.PartUri.Contains("footer", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(fields, field =>
                    field.LocationKind == WordFieldLocationKind.Footnote &&
                    field.FieldType == WordFieldType.Author);
                Assert.Contains(fields, field =>
                    field.LocationKind == WordFieldLocationKind.Endnote &&
                    field.FieldType == WordFieldType.Subject);
                Assert.Contains(fields, field =>
                    field.LocationKind == WordFieldLocationKind.Body &&
                    field.FieldType == WordFieldType.Keywords &&
                    field.IsInContentControl);
                Assert.Contains(fields, field =>
                    field.LocationKind == WordFieldLocationKind.Body &&
                    field.FieldType == WordFieldType.Subject &&
                    field.IsInTextBox);
            }
        }

        [Fact]
        public void Test_InspectFields_IgnoresExtensionFieldsAndUnselectedCompatibilityBranches() {
            const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            const string wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
            string path = Path.Combine(_directoryWithFiles, "FieldInventory.ExtensionMarkup.docx");

            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Normal story").AddField(WordFieldType.Author);
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(path, true)) {
                Paragraph paragraph = package.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
                paragraph.InnerXml +=
                    $"<x:extLst xmlns:x=\"urn:field-extension\" xmlns:w=\"{w}\"><w:fldSimple w:instr=\" TITLE \"/></x:extLst>" +
                    $"<mc:AlternateContent xmlns:mc=\"{mc}\" xmlns:w=\"{w}\">" +
                    "<mc:Choice xmlns:wps=\"urn:unsupported\" Requires=\"wps\"><w:fldSimple w:instr=\" SUBJECT \"/></mc:Choice>" +
                    $"<mc:Choice xmlns:shape=\"{wps}\" Requires=\"shape\"><w:fldSimple w:instr=\" FILENAME \"/></mc:Choice>" +
                    "<mc:Fallback><w:fldSimple w:instr=\" COMMENTS \"/></mc:Fallback></mc:AlternateContent>";
                package.MainDocumentPart.Document.Save();
            }

            using WordDocument loaded = WordDocument.Load(path);
            Assert.Equal(new[] { WordFieldType.Author, WordFieldType.FileName },
                loaded.InspectFields().Select(field => field.FieldType!.Value));
#if NET
            Assert.True(loaded.IsSdkDeferred);
#endif

            _ = loaded.OpenXmlDocument;
            Assert.Equal(new[] { WordFieldType.Author, WordFieldType.FileName },
                loaded.InspectFields().Select(field => field.FieldType!.Value));
        }

        [Fact]
        public void Test_DisposedAuthoredFieldResolvesWhenParagraphOpensSdkCopy() {
            WordField field;
            using (WordDocument document = WordDocument.Create()) {
                field = document.AddParagraph().AddField(WordFieldType.Page);
                field.Text = "7";
            }

            Assert.NotNull(field.Paragraph._paragraph.Parent);
            Assert.Equal("7", field.Text);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_AuthoredFieldEditsTheSameResultAfterOpeningTheSdk(bool advanced) {
            using WordDocument document = WordDocument.Create();
            var time = new DateTime(2025, 3, 17, 12, 34, 56);
            WordParagraph paragraph = document.AddParagraph("Value: ");
            WordField field = paragraph.AddField(WordFieldType.Time, customFormat: "HH:mm", advanced: advanced);

            Assert.Same(paragraph, field.Paragraph);
            Assert.Equal(WordFieldType.Time, field.FieldType);
            Assert.True(field.UpdateField);
            WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                CurrentDateTime = time,
                DateTimeFieldsOnly = true
            }).Results);
            Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
            Assert.Equal("12:34", field.Text);
#if NET
            Assert.True(document.IsSdkDeferred);
#endif

            _ = document.OpenXmlDocument;
            field.Text = "kept result";
            field.LockField = true;

            using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
            WordFieldInfo saved = Assert.Single(reopened.InspectFields());
            Assert.Equal("kept result", saved.ResultText);
            Assert.True(saved.IsLocked);
        }

        private static void AddSplitComplexField(Paragraph paragraph, string instructionPart1, string instructionPart2, string resultText, bool dirty, bool locked) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = dirty, FieldLock = locked }),
                new Run(new FieldCode { Text = instructionPart1, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldCode { Text = instructionPart2, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void AddNestedComplexFields(Paragraph paragraph) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = " QUOTE ", Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Outer start ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = " AUTHOR ", Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Nested Author") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(" outer end") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static SimpleField BuildSimpleField(string instruction, string resultText) {
            return new SimpleField(
                new Run(
                    new Text(resultText) { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = instruction
            };
        }

        private static void AddContentControlField(WordDocument document, string instruction, string resultText) {
            Body body = document._document.Body ?? throw new InvalidOperationException("Document body is missing.");
            var contentControl = new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = "FieldHost" },
                    new Tag { Val = "FieldHost" }),
                new SdtContentBlock(
                    new Paragraph(BuildSimpleField(instruction, resultText))));

            SectionProperties? sectionProperties = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectionProperties == null) {
                body.Append(contentControl);
            } else {
                body.InsertBefore(contentControl, sectionProperties);
            }
        }

        private static void AppendNoteFields(string filePath) {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true);
            MainDocumentPart mainPart = wordDocument.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");

            Footnote footnote = mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>().First(note => note.Type == null)
                ?? throw new InvalidOperationException("Test footnote is missing.");
            footnote.Append(new Paragraph(BuildSimpleField(" AUTHOR ", "Footnote Author")));
            mainPart.FootnotesPart!.Footnotes!.Save();

            Endnote endnote = mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>().First(note => note.Type == null)
                ?? throw new InvalidOperationException("Test endnote is missing.");
            endnote.Append(new Paragraph(BuildSimpleField(" SUBJECT ", "Endnote Subject")));
            mainPart.EndnotesPart!.Endnotes!.Save();
        }

        private static WordHeader RequireSectionHeader(WordDocument document, int index, HeaderFooterValues type) {
            Assert.NotNull(document);
            Assert.InRange(index, 0, document.Sections.Count - 1);

            var section = document.Sections[index];
            if (type == HeaderFooterValues.Default && section.PageSetup.GetHeader(WordHeaderFooterType.Default) == null) {
                _ = section.PageSetup.Header;
            }

            return type == HeaderFooterValues.Default
                ? Assert.IsAssignableFrom<WordHeader>(section.PageSetup.GetHeader(WordHeaderFooterType.Default))
                : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
        }

        private static WordFooter RequireSectionFooter(WordDocument document, int index, HeaderFooterValues type) {
            Assert.NotNull(document);
            Assert.InRange(index, 0, document.Sections.Count - 1);

            var section = document.Sections[index];
            if (type == HeaderFooterValues.Default && section.PageSetup.GetFooter(WordHeaderFooterType.Default) == null) {
                _ = section.PageSetup.Header;
            }

            return type == HeaderFooterValues.Default
                ? Assert.IsAssignableFrom<WordFooter>(section.PageSetup.GetFooter(WordHeaderFooterType.Default))
                : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");
        }
    }
}
