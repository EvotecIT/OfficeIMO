using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSimpleFieldBookmarksAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph bodyParagraph = document.AddParagraph("Body ");
                    AppendBookmarkedSimpleField(bodyParagraph._paragraph, "81", "BodySimpleFieldBookmark", " PAGE ");

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell ", removeExistingParagraphs: true);
                    AppendBookmarkedSimpleField(cellParagraph._paragraph, "82", "CellSimpleFieldBookmark", " NUMPAGES ");

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    WordParagraph headerParagraph = section.Header.Default!.AddParagraph("Header ");
                    AppendBookmarkedSimpleField(headerParagraph._paragraph, "83", "HeaderSimpleFieldBookmark", " PAGE ");

                    WordParagraph footerParagraph = section.Footer.Default!.AddParagraph("Footer ");
                    AppendBookmarkedSimpleField(footerParagraph._paragraph, "84", "FooterSimpleFieldBookmark", " NUMPAGES ");

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    AppendBookmarkedSimpleField(footnoteBody._paragraph, "85", "FootnoteSimpleFieldBookmark", " PAGE ");

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    AppendBookmarkedSimpleField(endnoteBody._paragraph, "86", "EndnoteSimpleFieldBookmark", " NUMPAGES ");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                AssertBookmarkRoundTrip(reloaded, "BodySimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "CellSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "HeaderSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FooterSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FootnoteSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "EndnoteSimpleFieldBookmark");
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocComplexFieldBookmarksAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph bodyParagraph = document.AddParagraph("Body ");
                    AppendBookmarkedComplexField(bodyParagraph._paragraph, "91", "BodyComplexFieldBookmark", " PAGE ");

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell ", removeExistingParagraphs: true);
                    AppendBookmarkedComplexField(cellParagraph._paragraph, "92", "CellComplexFieldBookmark", " NUMPAGES ");

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    WordParagraph headerParagraph = section.Header.Default!.AddParagraph("Header ");
                    AppendBookmarkedComplexField(headerParagraph._paragraph, "93", "HeaderComplexFieldBookmark", " PAGE ");

                    WordParagraph footerParagraph = section.Footer.Default!.AddParagraph("Footer ");
                    AppendBookmarkedComplexField(footerParagraph._paragraph, "94", "FooterComplexFieldBookmark", " NUMPAGES ");

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    AppendBookmarkedComplexField(footnoteBody._paragraph, "95", "FootnoteComplexFieldBookmark", " PAGE ");

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    AppendBookmarkedComplexField(endnoteBody._paragraph, "96", "EndnoteComplexFieldBookmark", " NUMPAGES ");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                AssertBookmarkRoundTrip(reloaded, "BodyComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "CellComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "HeaderComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FooterComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FootnoteComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "EndnoteComplexFieldBookmark");
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocDateFieldsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string dateInstruction = " DATE \\@ \"yyyy-MM-dd\" ";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph bodyParagraph = document.AddParagraph("Body date ");
                    AppendSimpleField(bodyParagraph._paragraph, dateInstruction, "2026-07-01");

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell date ", removeExistingParagraphs: true);
                    AppendComplexField(cellParagraph._paragraph, dateInstruction, "2026-07-02");

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    WordParagraph headerParagraph = section.Header.Default!.AddParagraph("Header date ");
                    AppendSimpleField(headerParagraph._paragraph, dateInstruction, "2026-07-03");

                    WordParagraph footerParagraph = section.Footer.Default!.AddParagraph("Footer date ");
                    AppendComplexField(footerParagraph._paragraph, dateInstruction, "2026-07-04");

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    AppendSimpleField(footnoteBody._paragraph, dateInstruction, "2026-07-05");

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    AppendComplexField(endnoteBody._paragraph, dateInstruction, "2026-07-06");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                MainDocumentPart mainPart = reloaded._wordprocessingDocument!.MainDocumentPart!;
                SimpleField[] dateFields = GetReloadedDateTimeFields(mainPart).ToArray();
                Assert.Equal(6, dateFields.Length);
                Assert.Contains(dateFields, field => IsFieldWithText(field, "DATE", "2026-07-01"));
                Assert.Contains(dateFields, field => IsFieldWithText(field, "DATE", "2026-07-02"));
                Assert.Contains(dateFields, field => IsFieldWithText(field, "DATE", "2026-07-03"));
                Assert.Contains(dateFields, field => IsFieldWithText(field, "DATE", "2026-07-04"));
                Assert.Contains(dateFields, field => IsFieldWithText(field, "DATE", "2026-07-05"));
                Assert.Contains(dateFields, field => IsFieldWithText(field, "DATE", "2026-07-06"));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocStaticDateTimeFieldsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    AppendSimpleField(document.AddParagraph("Time ")._paragraph, " TIME \\@ \"HH:mm\" ", "09:30");
                    AppendComplexField(document.AddParagraph("Created ")._paragraph, " CREATEDATE \\@ \"yyyy-MM-dd\" ", "2026-06-01");
                    AppendSimpleField(document.AddParagraph("Saved ")._paragraph, " SAVEDATE \\@ \"yyyy-MM-dd\" ", "2026-06-02");
                    AppendComplexField(document.AddParagraph("Printed ")._paragraph, " PRINTDATE \\@ \"yyyy-MM-dd\" ", "2026-06-03");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                SimpleField[] dateTimeFields = reloaded._wordprocessingDocument!.MainDocumentPart!.Document.Body!
                    .Descendants<SimpleField>()
                    .Where(IsStaticDateTimeField)
                    .ToArray();
                Assert.Equal(4, dateTimeFields.Length);
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "TIME", "09:30"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "CREATEDATE", "2026-06-01"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "SAVEDATE", "2026-06-02"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "PRINTDATE", "2026-06-03"));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocFieldResultInlineCharactersAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    AppendSimpleFieldWithResultRun(
                        document.AddParagraph("Body inline field ")._paragraph,
                        " DATE \\@ \"yyyy-MM-dd\" ",
                        "Body",
                        BreakValues.TextWrapping);

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell inline field ", removeExistingParagraphs: true);
                    AppendComplexFieldWithResultRun(
                        cellParagraph._paragraph,
                        " TIME \\@ \"HH:mm\" ",
                        "Cell",
                        BreakValues.Page);

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    AppendSimpleFieldWithResultRun(
                        section.Header.Default!.AddParagraph("Header inline field ")._paragraph,
                        " CREATEDATE \\@ \"yyyy-MM-dd\" ",
                        "Header",
                        BreakValues.Column);

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    AppendSimpleFieldWithResultRun(
                        footnoteBody._paragraph,
                        " SAVEDATE \\@ \"yyyy-MM-dd\" ",
                        "Footnote",
                        BreakValues.TextWrapping);

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    AppendComplexFieldWithResultRun(
                        endnoteBody._paragraph,
                        " PRINTDATE \\@ \"yyyy-MM-dd\" ",
                        "Endnote",
                        BreakValues.Page);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                SimpleField[] fields = GetReloadedDateTimeFields(reloaded._wordprocessingDocument!.MainDocumentPart!).ToArray();
                AssertFieldResultInlineContent(fields, "DATE", "BodyTextSoftHyphenHardHyphenWrap", BreakValues.TextWrapping);
                AssertFieldResultInlineContent(fields, "TIME", "CellTextSoftHyphenHardHyphenWrap", BreakValues.Page);
                AssertFieldResultInlineContent(fields, "CREATEDATE", "HeaderTextSoftHyphenHardHyphenWrap", BreakValues.Column);
                AssertFieldResultInlineContent(fields, "SAVEDATE", "FootnoteTextSoftHyphenHardHyphenWrap", BreakValues.TextWrapping);
                AssertFieldResultInlineContent(fields, "PRINTDATE", "EndnoteTextSoftHyphenHardHyphenWrap", BreakValues.Page);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocInlineContentControlStaticDateTimeFieldsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph bodyParagraph = document.AddParagraph("Body controlled ");
                    bodyParagraph._paragraph.Append(CreateInlineContentControl(
                        "Legacy DOC body inline field",
                        CreateSimpleField(" DATE \\@ \"yyyy-MM-dd\" ", "2026-07-01")));

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell controlled ", removeExistingParagraphs: true);
                    cellParagraph._paragraph.Append(CreateInlineContentControl(
                        "Legacy DOC table inline field",
                        CreateSimpleField(" TIME \\@ \"HH:mm\" ", "09:30")));

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    WordParagraph headerParagraph = section.Header.Default!.AddParagraph("Header controlled ");
                    headerParagraph._paragraph.Append(CreateInlineContentControl(
                        "Legacy DOC header inline field",
                        CreateSimpleField(" CREATEDATE \\@ \"yyyy-MM-dd\" ", "2026-06-01")));

                    WordParagraph footerParagraph = section.Footer.Default!.AddParagraph("Footer controlled ");
                    footerParagraph._paragraph.Append(CreateInlineContentControl(
                        "Legacy DOC footer inline field",
                        CreateSimpleField(" SAVEDATE \\@ \"yyyy-MM-dd\" ", "2026-06-02")));

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    footnoteBody._paragraph.Append(CreateInlineContentControl(
                        "Legacy DOC footnote inline field",
                        CreateSimpleField(" PRINTDATE \\@ \"yyyy-MM-dd\" ", "2026-06-03")));

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    endnoteBody._paragraph.Append(CreateInlineContentControl(
                        "Legacy DOC endnote inline field",
                        CreateSimpleField(" DATE \\@ \"yyyy-MM-dd\" ", "2026-07-02")));

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                MainDocumentPart mainPart = reloaded._wordprocessingDocument!.MainDocumentPart!;
                SimpleField[] dateTimeFields = GetReloadedDateTimeFields(mainPart).ToArray();
                Assert.Equal(6, dateTimeFields.Length);
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "DATE", "2026-07-01"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "TIME", "09:30"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "CREATEDATE", "2026-06-01"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "SAVEDATE", "2026-06-02"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "PRINTDATE", "2026-06-03"));
                Assert.Contains(dateTimeFields, field => IsFieldWithText(field, "DATE", "2026-07-02"));
                Assert.Empty(mainPart.Document.Descendants<SdtRun>());
                Assert.Empty(mainPart.HeaderParts.SelectMany(part => part.Header.Descendants<SdtRun>()));
                Assert.Empty(mainPart.FooterParts.SelectMany(part => part.Footer.Descendants<SdtRun>()));
                Assert.Empty(mainPart.FootnotesPart!.Footnotes!.Descendants<SdtRun>());
                Assert.Empty(mainPart.EndnotesPart!.Endnotes!.Descendants<SdtRun>());
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static void AppendBookmarkedSimpleField(Paragraph paragraph, string id, string name, string instruction) {
            var simpleField = new SimpleField { Instruction = instruction };
            simpleField.Append(
                new BookmarkStart { Id = id, Name = name },
                new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = id });
            paragraph.Append(simpleField);
        }

        private static void AppendBookmarkedComplexField(Paragraph paragraph, string id, string name, string instruction) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new BookmarkStart { Id = id, Name = name },
                new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = id },
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void AppendSimpleField(Paragraph paragraph, string instruction, string resultText) {
            var simpleField = new SimpleField { Instruction = instruction };
            simpleField.Append(new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.Append(simpleField);
        }

        private static void AppendComplexField(Paragraph paragraph, string instruction, string resultText) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void AppendSimpleFieldWithResultRun(Paragraph paragraph, string instruction, string prefix, BreakValues breakType) {
            var simpleField = new SimpleField { Instruction = instruction };
            simpleField.Append(CreateInlineFieldResultRun(prefix, breakType));
            paragraph.Append(simpleField);
        }

        private static void AppendComplexFieldWithResultRun(Paragraph paragraph, string instruction, string prefix, BreakValues breakType) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                CreateInlineFieldResultRun(prefix, breakType),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static Run CreateInlineFieldResultRun(string prefix, BreakValues breakType) {
            return new Run(
                new Text(prefix + "Text") { Space = SpaceProcessingModeValues.Preserve },
                new TabChar(),
                new SoftHyphen(),
                new Text("SoftHyphen") { Space = SpaceProcessingModeValues.Preserve },
                new NoBreakHyphen(),
                new Text("HardHyphen") { Space = SpaceProcessingModeValues.Preserve },
                breakType == BreakValues.TextWrapping ? new Break() : new Break { Type = breakType },
                new Text("Wrap") { Space = SpaceProcessingModeValues.Preserve });
        }

        private static SimpleField CreateSimpleField(string instruction, string resultText) {
            var simpleField = new SimpleField { Instruction = instruction };
            simpleField.Append(CreateTextRun(resultText));
            return simpleField;
        }

        private static IEnumerable<SimpleField> GetReloadedDateTimeFields(MainDocumentPart mainPart) {
            foreach (SimpleField field in mainPart.Document.Body!.Descendants<SimpleField>()) {
                if (IsStaticDateTimeField(field)) {
                    yield return field;
                }
            }

            foreach (SimpleField field in mainPart.HeaderParts.SelectMany(part => part.Header.Descendants<SimpleField>())) {
                if (IsStaticDateTimeField(field)) {
                    yield return field;
                }
            }

            foreach (SimpleField field in mainPart.FooterParts.SelectMany(part => part.Footer.Descendants<SimpleField>())) {
                if (IsStaticDateTimeField(field)) {
                    yield return field;
                }
            }

            if (mainPart.FootnotesPart?.Footnotes != null) {
                foreach (SimpleField field in mainPart.FootnotesPart.Footnotes.Descendants<SimpleField>()) {
                    if (IsStaticDateTimeField(field)) {
                        yield return field;
                    }
                }
            }

            if (mainPart.EndnotesPart?.Endnotes != null) {
                foreach (SimpleField field in mainPart.EndnotesPart.Endnotes.Descendants<SimpleField>()) {
                    if (IsStaticDateTimeField(field)) {
                        yield return field;
                    }
                }
            }
        }

        private static bool IsFieldWithText(SimpleField field, string fieldName, string expectedText) {
            return IsFieldInstruction(field, fieldName)
                && string.Concat(field.Descendants<Text>().Select(text => text.Text)) == expectedText;
        }

        private static void AssertFieldResultInlineContent(IEnumerable<SimpleField> fields, string fieldName, string expectedText, BreakValues expectedBreakType) {
            SimpleField field = Assert.Single(fields, field => IsFieldWithText(field, fieldName, expectedText));
            Assert.Single(field.Descendants<TabChar>());
            Assert.NotEmpty(field.Descendants<SoftHyphen>());
            Assert.NotEmpty(field.Descendants<NoBreakHyphen>());
            Break resultBreak = Assert.Single(field.Descendants<Break>());
            if (expectedBreakType == BreakValues.TextWrapping) {
                Assert.True(resultBreak.Type == null || resultBreak.Type.Value == BreakValues.TextWrapping);
            } else {
                Assert.Equal(expectedBreakType, resultBreak.Type?.Value);
            }
        }

        private static bool IsStaticDateTimeField(SimpleField field) {
            return IsFieldInstruction(field, "DATE")
                || IsFieldInstruction(field, "TIME")
                || IsFieldInstruction(field, "CREATEDATE")
                || IsFieldInstruction(field, "SAVEDATE")
                || IsFieldInstruction(field, "PRINTDATE");
        }

        private static bool IsFieldInstruction(SimpleField field, string fieldName) {
            string instruction = field.Instruction?.Value ?? string.Empty;
            string trimmed = instruction.TrimStart();
            return trimmed.StartsWith(fieldName, StringComparison.OrdinalIgnoreCase)
                && (trimmed.Length == fieldName.Length || char.IsWhiteSpace(trimmed[fieldName.Length]));
        }
    }
}
