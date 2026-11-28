using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private const string TocFieldEndParagraphFixture = "word-generated-toc-field-end-paragraph.docx";
        private const string TocContentControlFixture = "word-generated-toc-gallery-pages.docx";

        [Theory]
        [InlineData(TocFieldEndParagraphFixture)]
        [InlineData(TocContentControlFixture)]
        public void Test_TocRegenerationKeepsEveryFieldEndAfterItsBegin(string fixture) {
            using WordDocument document = LoadTocFixture(fixture);

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());

            AssertFieldCharactersNested(document._document.Body!);
            SdtContentBlock content = document.TableOfContent!.SdtBlock.SdtContentBlock!;
            // Word's ToC begins with its first entry (or the content control's title); there is no empty line above it.
            Paragraph first = content.Elements<Paragraph>().First(paragraph => !IsTocHeadingParagraph(paragraph));
            Assert.Contains(first.Descendants<FieldCode>(), code => code.Text.TrimStart().StartsWith("TOC", StringComparison.Ordinal));
            var errors = document.ValidateDocument();
            Assert.True(errors.Count == 0, FormatValidationErrors(errors));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_TocWithoutTitleGainsNoTitleWhenRegenerated(bool splitTocKeyword) {
            using WordDocument document = LoadTocFixture(TocFieldEndParagraphFixture);
            if (splitTocKeyword) {
                // Word can split an instruction across runs, even inside the TOC keyword.
                FieldCode code = document.TableOfContent!.SdtBlock.Descendants<FieldCode>()
                    .First(code => code.Text.TrimStart().StartsWith("TOC", StringComparison.Ordinal));
                int keyword = code.Text.IndexOf("TOC", StringComparison.Ordinal);
                code.Parent!.InsertAfterSelf(new Run(new FieldCode(code.Text.Substring(keyword + 1)) { Space = SpaceProcessingModeValues.Preserve }));
                code.Text = code.Text.Substring(0, keyword + 1);
                code.Space = SpaceProcessingModeValues.Preserve;
            }

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());

            Paragraph first = document.TableOfContent!.SdtBlock.SdtContentBlock!.Elements<Paragraph>().First();
            Assert.StartsWith("Topic 1", string.Concat(first.Descendants<Text>().Select(text => text.Text)), StringComparison.Ordinal);
            Assert.Equal(30, document.TableOfContent.SdtBlock.Descendants<Hyperlink>().Count());
        }

        [Fact]
        public void Test_TocTitleWithItsOwnFieldIsKeptWhenRegenerated() {
            using WordDocument document = LoadTocFixture(TocFieldEndParagraphFixture);
            SdtContentBlock content = document.TableOfContent!.SdtBlock.SdtContentBlock!;
            content.InsertAt(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(new Text("Contents of ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" DOCPROPERTY Title ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Report")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })), 0);

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());

            Paragraph first = document.TableOfContent!.SdtBlock.SdtContentBlock!.Elements<Paragraph>().First();
            Assert.StartsWith("Contents of ", string.Concat(first.Descendants<Text>().Select(text => text.Text)), StringComparison.Ordinal);
            Assert.Contains(first.Descendants<FieldCode>(), code => code.Text.Contains("DOCPROPERTY", StringComparison.Ordinal));
            Assert.Equal(30, document.TableOfContent.SdtBlock.Descendants<Hyperlink>().Count());
            AssertFieldCharactersNested(document._document.Body!);
        }

        [Fact]
        public void Test_TocWithoutTitleIsDrawnWithoutTitleInPdf() {
            using WordDocument document = CreateLeveledHeadings();
            SdtContentBlock content = document.TableOfContent!.SdtBlock.SdtContentBlock!;
            foreach (Paragraph title in content.Elements<Paragraph>().Where(IsTocHeadingParagraph).ToList()) title.Remove();
            Assert.Null(document.TableOfContent.TitleParagraph);

            using var stream = new MemoryStream();
            document.SaveAsPdf(stream);
            string firstPage = PdfDocument.Load(stream.ToArray()).Reader.TextByPage()[0];

            Assert.Contains("Entry 1", firstPage, StringComparison.Ordinal);
            Assert.DoesNotContain("Table of Contents", firstPage, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_GeneratedTocDropsContextualSpacingBetweenEntriesOfOneLevel() {
            using WordDocument generated = CreateLeveledHeadings();
            using WordDocument stored = CreateLeveledHeadings();
            foreach (WordDocument document in new[] { generated, stored }) {
                Style entryStyle = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                    .Elements<Style>().Single(style => style.StyleId?.Value == "TOC2");
                StyleParagraphProperties properties = entryStyle.StyleParagraphProperties ??= new StyleParagraphProperties();
                properties.SpacingBetweenLines = new SpacingBetweenLines { Before = "120", After = "360" };
                properties.ContextualSpacing = new ContextualSpacing();
                foreach (WordParagraph paragraph in document.Paragraphs.Where(paragraph => paragraph.Style == WordParagraphStyles.Heading3)) {
                    paragraph.Style = WordParagraphStyles.Heading2;
                }
            }
            stored.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
            stored.Settings.UpdateFieldsOnOpen = false;

            // Entries 2-3, 5-6, 8-9 and 11-12 are consecutive level-2 entries: Word drops the space between them.
            // Between levels, the space before a level-2 entry is shared with the level-1 entry's space after.
            IReadOnlyList<(string Text, double X, double Y)> generatedEntries = ReadEntryPositions(generated);
            IReadOnlyList<(string Text, double X, double Y)> storedEntries = ReadEntryPositions(stored);

            Assert.Equal(12, storedEntries.Count);
            for (int index = 0; index < storedEntries.Count; index++) Assert.Equal(storedEntries[index].Y, generatedEntries[index].Y, 1);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_BareIndexWrapsAndRefreshesWithTheOwnedPackage(bool compatibilitySuffix) {
            byte[] bytes;
            using (WordDocument source = WordDocument.Create()) {
                Paragraph paragraph = source.AddParagraph("Text before index")._paragraph;
                paragraph.Append(new SimpleField(new Run(new Text("Stale index"))) { Instruction = " INDEX " });
                var suffix = new Run(new Text("Text after index"));
                if (compatibilitySuffix) {
                    paragraph.Append(new AlternateContent(
                        new AlternateContentChoice(new Run()) { Requires = "w14" },
                        new AlternateContentFallback(suffix)));
                } else {
                    paragraph.Append(suffix);
                }
                source.AddParagraph("Indexed text")._paragraph.Append(new SimpleField() { Instruction = " XE Known " });
                bytes = source.ToBytes();
            }

            using WordDocument document = WordDocument.Load(new MemoryStream(bytes));
            WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
            Assert.Equal("Known", Assert.Single(index.RefreshIndex().Entries).Term);
#if NET
            Assert.True(document.IsSdkDeferred);
#endif

            var body = document.OpenXmlDocument.MainDocumentPart!.Document.Body!;
            Paragraph[] paragraphs = body.Elements<Paragraph>().ToArray();
            Assert.Equal(3, paragraphs.Length);
            Assert.Equal("Text before index", paragraphs[0].InnerText);
            if (compatibilitySuffix) {
                var alternate = Assert.Single(paragraphs[1].Elements<AlternateContent>());
                Assert.Equal("Text after index", alternate.GetFirstChild<AlternateContentFallback>()!.InnerText);
            } else {
                Assert.Equal("Text after index", paragraphs[1].InnerText);
            }
            Assert.Equal("Indexed text", paragraphs[2].InnerText);
            Assert.Single(body.Elements<SdtBlock>());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_CaptionAndIndexRefreshIgnoreExtensionFields(bool captionList) {
            const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            string instruction = captionList ? " SEQ Figure " : " XE Known ";
            string hiddenInstruction = captionList ? " SEQ Figure " : " XE Hidden ";
            byte[] bytes;
            using (WordDocument source = WordDocument.Create()) {
                source.AddTableOfContent();
                source.AddParagraph("Visible source")._paragraph.Append(
                    new SimpleField(new Run(new Text("1"))) { Instruction = instruction });
                var extension = new OpenXmlUnknownElement("w", "extLst", w);
                extension.Append(new SimpleField(new Run(new Text("99"))) { Instruction = hiddenInstruction });
                source.AddParagraph("Extension source")._paragraph.Append(extension);
                bytes = source.ToBytes();
            }

            using WordDocument document = WordDocument.Load(new MemoryStream(bytes));
            WordTableOfContent table = Assert.IsType<WordTableOfContent>(document.TableOfContent);
            if (captionList) {
                Assert.Single(table.RefreshListOfFigures().Entries);
            } else {
                Assert.Equal("Known", Assert.Single(table.RefreshIndex().Entries).Term);
            }
#if NET
            Assert.True(document.IsSdkDeferred);
#endif
        }

        [Theory]
        [InlineData("figures")]
        [InlineData("index")]
        public void Test_CaptionListAndIndexRefreshKeepFieldCharactersNested(string kind) {
            using WordDocument document = LoadTocFixture(TocFieldEndParagraphFixture);
            WordTableOfContent table = document.TableOfContent!;

            if (kind == "figures") table.RefreshListOfFigures();
            else table.RefreshIndex();

            AssertFieldCharactersNested(document._document.Body!);
            var errors = document.ValidateDocument();
            Assert.True(errors.Count == 0, FormatValidationErrors(errors));
        }

        [Theory]
        [InlineData(TocFieldEndParagraphFixture)]
        [InlineData(TocContentControlFixture)]
        public void Test_WordTocFixturesKeepWordPagesInPdfAndFieldPagination(string fixture) {
            // Word lays out both documents on 6 pages: 28 ToC entries fit page 1, so Topic 1 starts on page 3.
            // Fonts with Aptos's metrics give the same layout on hosts without Aptos.
            WordPdfExportOptions options = CreateAptosMetricOptions();
            using (WordDocument document = LoadTocFixture(fixture)) {
                Assert.Equal(new TocPdfPages(6, 3, 3), ReadTocPdfPages(document, options));
                using var stream = new MemoryStream();
                document.SaveAsPdf(stream, options);
                using var pdf = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
                Assert.All(pdf.GetPages().SelectMany(page => page.Letters).Where(letter => !char.IsWhiteSpace(letter.Value[0])),
                    letter => Assert.Contains("Metrics", letter.FontName, StringComparison.Ordinal));
            }

            using (WordDocument document = LoadTocFixture(fixture)) {
                document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination(options));
                Assert.Equal(new TocPdfPages(6, 3, 3), ReadTocPdfPages(document, options));
                // A second update starts from the content control OfficeIMO wrote.
                document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination(options));
                Assert.Equal(new TocPdfPages(6, 3, 3), ReadTocPdfPages(document, options));
            }
        }

        [Fact]
        public void Test_GeneratedTocUsesTheEntryStylesOfTheStoredToc() {
            using WordDocument generated = CreateLeveledHeadings();
            using WordDocument stored = CreateLeveledHeadings();
            stored.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
            stored.Settings.UpdateFieldsOnOpen = false;

            // The placeholder table is drawn from the headings; the regenerated one as stored. Both use TOC 1-3.
            IReadOnlyList<(string Text, double X, double Y)> generatedEntries = ReadEntryPositions(generated);
            IReadOnlyList<(string Text, double X, double Y)> storedEntries = ReadEntryPositions(stored);

            Assert.Equal(12, storedEntries.Count);
            Assert.Equal(storedEntries.Select(entry => entry.Text), generatedEntries.Select(entry => entry.Text));
            for (int index = 0; index < storedEntries.Count; index++) {
                Assert.Equal(storedEntries[index].X, generatedEntries[index].X, 1);
                Assert.Equal(storedEntries[index].Y, generatedEntries[index].Y, 1);
            }
        }

        private WordDocument LoadTocFixture(string fixture) {
            string path = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", fixture));
            using var stream = new MemoryStream(File.ReadAllBytes(path));
            return WordDocument.Load(stream);
        }

        private static WordDocument CreateLeveledHeadings() {
            WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddPageBreak();
            for (int index = 1; index <= 12; index++) {
                WordParagraphStyles style = (index % 3) switch {
                    1 => WordParagraphStyles.Heading1,
                    2 => WordParagraphStyles.Heading2,
                    _ => WordParagraphStyles.Heading3
                };
                document.AddParagraph("Entry " + index).Style = style;
                document.AddParagraph("Body " + index);
            }
            return document;
        }

        private static bool IsTocHeadingParagraph(Paragraph paragraph) =>
            !paragraph.Descendants<FieldChar>().Any() && !paragraph.Descendants<Hyperlink>().Any() && paragraph.InnerText.Length > 0;

        private static void AssertFieldCharactersNested(OpenXmlElement root) {
            int depth = 0;
            foreach (FieldChar fieldChar in root.Descendants<FieldChar>()) {
                if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin) depth++;
                else if (fieldChar.FieldCharType?.Value == FieldCharValues.End) {
                    depth--;
                    Assert.True(depth >= 0, "A field end comes before its begin.");
                } else if (fieldChar.FieldCharType?.Value == FieldCharValues.Separate) {
                    Assert.True(depth > 0, "A field separator is outside a field.");
                }
            }
            Assert.Equal(0, depth);
        }

        // Advances of the printable ASCII characters (32-126) in Aptos and Aptos Display, 2048 units per em.
        private const string AptosAdvances = "416,600,757,1100,1094,1692,1317,431,600,600,936,1094,585,697,585,695,1094,1094,1094,1094,1094,1094,1094,1094,1094,1094,585,585,1094,1094,1094,1027,1832,1207,1237,1418,1405,1139,1074,1451,1448,533,678,1164,1023,1618,1446,1500,1182,1500,1241,1159,980,1395,1199,1827,1132,1105,1056,603,695,603,1094,942,1130,1088,1149,1075,1148,1079,617,992,1129,489,489,997,533,1747,1129,1130,1149,1148,685,995,662,1144,926,1476,906,926,897,603,553,603,1094";
        private const string AptosDisplayAdvances = "395,504,757,1082,1071,1690,1276,431,603,603,780,1071,504,714,504,695,1071,1071,1071,1071,1071,1071,1071,1071,1071,1071,504,504,1071,1071,1071,996,1799,1188,1193,1343,1341,1107,1043,1375,1355,481,650,1134,992,1567,1394,1424,1138,1424,1209,1118,977,1333,1175,1814,1127,1095,1056,603,695,603,1071,942,1130,1002,1062,1009,1062,1012,580,900,1036,431,431,916,499,1638,1036,1050,1062,1062,635,948,622,1036,886,1466,902,886,872,603,449,603,1071";

        /// <summary>Export options that lay Aptos text out with test fonts carrying Aptos's metrics and no installed fonts.</summary>
        private static WordPdfExportOptions CreateAptosMetricOptions() {
            var pdf = new PdfCreateOptions();
            foreach ((string family, string advances) in new[] { ("Aptos", AptosAdvances), ("Aptos Display", AptosDisplayAdvances) }) {
                ushort[] widths = advances.Split(',').Select(value => ushort.Parse(value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                Dictionary<int, ushort> byCharacter = Enumerable.Range(0, widths.Length).ToDictionary(index => 32 + index, index => widths[index]);
                string metricFamily = "OfficeIMO " + family + " Metrics";
                pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(metricFamily,
                    OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithLayoutMetrics(2048, 1923, -577, 2068, 563, byCharacter)));
                pdf.RegisterFontFamilySubstitution(family, metricFamily, PdfFontFamilySubstitutionImpact.Compatible);
            }
            return new WordPdfExportOptions { PdfOptions = pdf, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        }

        private readonly record struct TocPdfPages(int PageCount, int TocTopicOnePage, int TopicOnePage);

        private static TocPdfPages ReadTocPdfPages(WordDocument document, WordPdfExportOptions options) {
            using var stream = new MemoryStream();
            document.SaveAsPdf(stream, options);
            IReadOnlyList<string> pages = PdfDocument.Load(stream.ToArray()).Reader.TextByPage();
            var entry = System.Text.RegularExpressions.Regex.Match(pages[0], @"Topic 1\b[\s.]*(\d+)");
            Assert.True(entry.Success, pages[0]);
            int topicOne = Enumerable.Range(0, pages.Count).First(index => pages[index].Contains("Body line for topic 1.", StringComparison.Ordinal)) + 1;
            return new TocPdfPages(pages.Count, int.Parse(entry.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture), topicOne);
        }

        private static IReadOnlyList<(string Text, double X, double Y)> ReadEntryPositions(WordDocument document) {
            using var stream = new MemoryStream();
            document.SaveAsPdf(stream);
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
            return pdf.GetPage(1).GetWords()
                .Where(word => word.Text == "Entry")
                .Select(word => ("Entry", Math.Round(word.BoundingBox.Left, 1), Math.Round(word.BoundingBox.Bottom, 1)))
                .ToList();
        }
    }
}
