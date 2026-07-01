using System;
using System.IO;
using System.Linq;
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
                document.Save(false);
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
        public void Test_InspectFields_ReadsFieldsAcrossPartsAndContainers() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldInventory.PartsAndContainers.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].AddField(WordFieldType.Page);

                document.AddHeadersAndFooters();
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
                document.Save(false);
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
            if (type == HeaderFooterValues.Default && section.Header.Default == null) {
                section.AddHeadersAndFooters();
            }

            return type == HeaderFooterValues.Default
                ? Assert.IsAssignableFrom<WordHeader>(section.Header.Default)
                : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
        }

        private static WordFooter RequireSectionFooter(WordDocument document, int index, HeaderFooterValues type) {
            Assert.NotNull(document);
            Assert.InRange(index, 0, document.Sections.Count - 1);

            var section = document.Sections[index];
            if (type == HeaderFooterValues.Default && section.Footer.Default == null) {
                section.AddHeadersAndFooters();
            }

            return type == HeaderFooterValues.Default
                ? Assert.IsAssignableFrom<WordFooter>(section.Footer.Default)
                : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");
        }
    }
}
