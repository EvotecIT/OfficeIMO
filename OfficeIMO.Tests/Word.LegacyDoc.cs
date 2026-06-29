using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenMcdf;
using Xunit;
using Version = OpenMcdf.Version;
using StorageModeFlags = OpenMcdf.StorageModeFlags;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsPlainTextParagraphs() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDoc("First paragraph", "Second paragraph");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(2, result.ImportReport.ParagraphCount);
            Assert.Equal(2, result.Document.Paragraphs.Count);
            Assert.Equal("First paragraph", result.Document.Paragraphs[0].Text);
            Assert.Equal("Second paragraph", result.Document.Paragraphs[1].Text);
            Assert.True(result.Document.WasLoadedFromLegacyDoc);
            Assert.Equal(string.Empty, result.Document.FilePath);

            using WordDocument reloaded = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            Assert.Equal("First paragraph", reloaded.Paragraphs[0].Text);
            Assert.Equal("Second paragraph", reloaded.Paragraphs[1].Text);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsDocumentPropertiesAndCustomProperties() {
            DateTime created = new DateTime(2026, 6, 29, 8, 0, 0, DateTimeKind.Utc);
            DateTime modified = new DateTime(2026, 6, 29, 9, 15, 0, DateTimeKind.Utc);
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithDocumentProperties(created, modified, "Metadata paragraph");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(13, result.ImportReport.DocumentPropertyCount);
            Assert.Equal("Legacy DOC Metadata Title", result.Document.BuiltinDocumentProperties.Title);
            Assert.Equal("Legacy DOC metadata subject", result.Document.BuiltinDocumentProperties.Subject);
            Assert.Equal("OfficeIMO Legacy Import", result.Document.BuiltinDocumentProperties.Creator);
            Assert.Equal("doc, metadata, officeimo", result.Document.BuiltinDocumentProperties.Keywords);
            Assert.Equal("OLE SummaryInformation comments", result.Document.BuiltinDocumentProperties.Description);
            Assert.Equal("Legacy Category", result.Document.BuiltinDocumentProperties.Category);
            AssertSameInstant(created, result.Document.BuiltinDocumentProperties.Created);
            AssertSameInstant(modified, result.Document.BuiltinDocumentProperties.Modified);
            Assert.Equal("EvotecIT", result.Document.ApplicationProperties.Company);
            Assert.Equal("Document Manager", result.Document.ApplicationProperties.Manager?.Text);
            Assert.Equal("Ready", result.Document.CustomDocumentProperties["ReleaseStatus"].Text);
            Assert.True(result.Document.CustomDocumentProperties["Reviewed"].Bool);
            Assert.Equal(2003, result.Document.CustomDocumentProperties["Ticket"].NumberInteger);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            Assert.False(converted.WasLoadedFromLegacyDoc);
            Assert.Equal("Legacy DOC Metadata Title", converted.BuiltinDocumentProperties.Title);
            Assert.Equal("EvotecIT", converted.ApplicationProperties.Company);
            Assert.Equal("Ready", converted.CustomDocumentProperties["ReleaseStatus"].Text);
            Assert.True(converted.CustomDocumentProperties["Reviewed"].Bool);
            Assert.Equal(2003, converted.CustomDocumentProperties["Ticket"].NumberInteger);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsDirectBoldItalicRuns() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithDirectCharacterFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] runs = result.Document.Paragraphs.ToArray();
            Assert.Equal(3, runs.Length);
            Assert.Equal("plain ", runs[0].Text);
            Assert.False(runs[0].Bold);
            Assert.False(runs[0].Italic);
            Assert.Equal("bold ", runs[1].Text);
            Assert.True(runs[1].Bold);
            Assert.False(runs[1].Italic);
            Assert.Equal("italic", runs[2].Text);
            Assert.False(runs[2].Bold);
            Assert.True(runs[2].Italic);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsDirectUnderlineSizeAndColorRuns() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithExtendedDirectCharacterFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] runs = result.Document.Paragraphs.ToArray();
            Assert.Equal(5, runs.Length);
            Assert.Equal("plain ", runs[0].Text);
            Assert.Null(runs[0].Underline);
            Assert.Equal("under ", runs[1].Text);
            Assert.Equal(UnderlineValues.Single, runs[1].Underline);
            Assert.Equal("sized ", runs[2].Text);
            Assert.Equal(14, runs[2].FontSize);
            Assert.Equal("red ", runs[3].Text);
            Assert.Equal("ff0000", runs[3].ColorHex);
            Assert.Equal("direct", runs[4].Text);
            Assert.Equal("336699", runs[4].ColorHex);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsFontFamilyRunsThroughFontTable() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithFontFamilyFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] runs = result.Document.Paragraphs.ToArray();
            Assert.Equal(2, runs.Length);
            Assert.Equal("plain ", runs[0].Text);
            Assert.Null(runs[0].FontFamily);
            Assert.Equal("font", runs[1].Text);
            Assert.Equal("Courier New", runs[1].FontFamily);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsParagraphAlignment() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithParagraphAlignment();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs.ToArray();
            Assert.Equal(3, paragraphs.Length);
            Assert.Equal("left", paragraphs[0].Text);
            Assert.Null(paragraphs[0].ParagraphAlignment);
            Assert.Equal("center", paragraphs[1].Text);
            Assert.Equal(JustificationValues.Center, paragraphs[1].ParagraphAlignment);
            Assert.Equal("right", paragraphs[2].Text);
            Assert.Equal(JustificationValues.Right, paragraphs[2].ParagraphAlignment);
        }

        [Fact]
        public void LegacyDoc_NormalLoad_RoutesOleDocIntoProjectedWordDocument() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDoc("Normal load"));

                using WordDocument document = WordDocument.Load(docPath);

                Assert.True(document.WasLoadedFromLegacyDoc);
                Assert.Equal(string.Empty, document.FilePath);
                WordParagraph paragraph = Assert.Single(document.Paragraphs);
                Assert.Equal("Normal load", paragraph.Text);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ImportsWordComDocFixture() {
            string docPath = GetFixtureDoc(Path.Combine("LegacyDocCorpus", "ComSimpleParagraphs.doc"));

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(docPath);

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.True(result.Document.WasLoadedFromLegacyDoc);
            Assert.Equal(string.Empty, result.Document.FilePath);

            string[] paragraphs = result.Document.Paragraphs
                .Select(paragraph => paragraph.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToArray();

            Assert.Contains("First COM paragraph", paragraphs);
            Assert.Contains("Second COM paragraph", paragraphs);
        }

        [Fact]
        public void LegacyDoc_CorpusImportReports_MatchCheckedInBaselines() {
            string corpusDirectory = Path.Combine(GetWordTestsProjectRoot(), "Documents", "LegacyDocCorpus");
            string[] docPaths = Directory.GetFiles(corpusDirectory, "*.doc", SearchOption.AllDirectories)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            Assert.NotEmpty(docPaths);

            bool updateBaselines = string.Equals(
                Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES"),
                "1",
                StringComparison.Ordinal);
            var missingBaselines = new List<string>();
            foreach (string docPath in docPaths) {
                using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(docPath);
                string actual = NormalizeLegacyDocBaselineText(result.ImportReport.ToMarkdown());
                string baselinePath = Path.ChangeExtension(docPath, ".import-report.md");

                if (updateBaselines) {
                    File.WriteAllText(baselinePath, actual, Encoding.UTF8);
                    continue;
                }

                if (!File.Exists(baselinePath)) {
                    missingBaselines.Add(Path.GetRelativePath(corpusDirectory, baselinePath));
                    continue;
                }

                string expected = NormalizeLegacyDocBaselineText(File.ReadAllText(baselinePath, Encoding.UTF8));
                Assert.Equal(expected, actual);
            }

            Assert.True(
                missingBaselines.Count == 0,
                "Missing legacy DOC corpus baselines. Run with OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES=1 to create: "
                    + string.Join(", ", missingBaselines));
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsMissingWordDocumentStream() {
            byte[] docBytes = LegacyDocTestBuilder.CreateCompoundWithoutWordDocumentStream();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            Assert.False(result.HasDocument);
            Assert.True(result.HasImportErrors);
            LegacyDocImportDiagnostic diagnostic = Assert.Single(result.Diagnostics);
            Assert.Equal("DOC-WORDDOCUMENT-MISSING", diagnostic.Code);
            Assert.Equal(LegacyDocDiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedCompoundFeatures() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Preserve-only body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(2, result.UnsupportedFeatures.Count);
            Assert.Equal(2, result.ImportReport.UnsupportedFeatureCount);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.OleObject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-MACROS-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-OLE-OBJECTS-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["VbaProject|DOC-MACROS-PRESENT|Compound:VbaProjectStorage"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["OleObject|DOC-OLE-OBJECTS-PRESENT|Compound:OleObjectStorage"]);
            Assert.Contains(result.UnsupportedFeatures, feature => feature.EntryPath == "_VBA_PROJECT_CUR");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.EntryPath == "ObjectPool");

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| Unsupported features | 2 |", markdown);
            Assert.Contains("| VbaProject | DOC-MACROS-PRESENT | Compound:VbaProjectStorage | _VBA_PROJECT_CUR |", markdown);
            Assert.Contains("| OleObject | DOC-OLE-OBJECTS-PRESENT | Compound:OleObjectStorage | ObjectPool |", markdown);
        }

        [Fact]
        public void LegacyDoc_NormalLoad_ExposesUnsupportedCompoundFeaturesOnProjectedDocument() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Normal load with unsupported features");

            using WordDocument document = WordDocument.Load(new MemoryStream(docBytes));

            Assert.True(document.WasLoadedFromLegacyDoc);
            Assert.Equal(2, document.LegacyDocUnsupportedFeatures.Count);
            Assert.Contains(document.LegacyDocUnsupportedFeatures, feature => feature.Kind == LegacyDocUnsupportedFeatureKind.VbaProject);
            Assert.Contains(document.LegacyDocUnsupportedFeatures, feature => feature.Kind == LegacyDocUnsupportedFeatureKind.OleObject);
            Assert.Contains(document.LegacyDocImportDiagnostics, diagnostic => diagnostic.Code == "DOC-MACROS-PRESENT");
            Assert.Contains(document.LegacyDocImportDiagnostics, diagnostic => diagnostic.Code == "DOC-OLE-OBJECTS-PRESENT");
        }

        [Fact]
        public void LegacyDoc_NormalLoad_BlocksAutoSaveForLegacyDocProjection() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDoc("No autosave"));

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => WordDocument.Load(docPath, autoSave: true));

                Assert.Contains("Auto-save is not supported", exception.Message);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Zażółć gęślą jaźń");
                    document.AddParagraph("Second plain paragraph");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(string.Empty, reloaded.FilePath);
                string[] paragraphs = reloaded.Paragraphs
                    .Select(paragraph => paragraph.Text)
                    .Where(text => !string.IsNullOrEmpty(text))
                    .ToArray();
                Assert.Equal(new[] { "Zażółć gęślą jaźń", "Second plain paragraph" }, paragraphs);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocPropertiesAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            DateTime created = new DateTime(2026, 6, 29, 10, 0, 0, DateTimeKind.Utc);
            DateTime modified = new DateTime(2026, 6, 29, 10, 30, 0, DateTimeKind.Utc);
            DateTime reviewedAt = new DateTime(2026, 6, 29, 11, 0, 0, DateTimeKind.Utc);

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Metadata native DOC");
                    document.BuiltinDocumentProperties.Title = "Native DOC Metadata Title";
                    document.BuiltinDocumentProperties.Subject = "Native DOC metadata subject";
                    document.BuiltinDocumentProperties.Creator = "OfficeIMO Native DOC";
                    document.BuiltinDocumentProperties.Keywords = "doc, metadata, native";
                    document.BuiltinDocumentProperties.Description = "Native DOC metadata comments";
                    document.BuiltinDocumentProperties.Category = "Native Category";
                    document.BuiltinDocumentProperties.Created = created;
                    document.BuiltinDocumentProperties.Modified = modified;
                    document.ApplicationProperties.Company = "EvotecIT";
                    document.ApplicationProperties.Manager = new Manager { Text = "Native Manager" };
                    document.CustomDocumentProperties["ReleaseStatus"] = new WordCustomProperty("Ready");
                    document.CustomDocumentProperties["Reviewed"] = new WordCustomProperty(true);
                    document.CustomDocumentProperties["Ticket"] = new WordCustomProperty(2004);
                    document.CustomDocumentProperties["Score"] = new WordCustomProperty(98.5d);
                    document.CustomDocumentProperties["ReviewedAt"] = new WordCustomProperty(reviewedAt);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal("Native DOC Metadata Title", reloaded.BuiltinDocumentProperties.Title);
                Assert.Equal("Native DOC metadata subject", reloaded.BuiltinDocumentProperties.Subject);
                Assert.Equal("OfficeIMO Native DOC", reloaded.BuiltinDocumentProperties.Creator);
                Assert.Equal("doc, metadata, native", reloaded.BuiltinDocumentProperties.Keywords);
                Assert.Equal("Native DOC metadata comments", reloaded.BuiltinDocumentProperties.Description);
                Assert.Equal("Native Category", reloaded.BuiltinDocumentProperties.Category);
                AssertSameInstant(created, reloaded.BuiltinDocumentProperties.Created);
                AssertSameInstant(modified, reloaded.BuiltinDocumentProperties.Modified);
                Assert.Equal("EvotecIT", reloaded.ApplicationProperties.Company);
                Assert.Equal("Native Manager", reloaded.ApplicationProperties.Manager?.Text);
                Assert.Equal("Ready", reloaded.CustomDocumentProperties["ReleaseStatus"].Text);
                Assert.True(reloaded.CustomDocumentProperties["Reviewed"].Bool);
                Assert.Equal(2004, reloaded.CustomDocumentProperties["Ticket"].NumberInteger);
                Assert.Equal(98.5d, reloaded.CustomDocumentProperties["Score"].NumberDouble);
                AssertSameInstant(reviewedAt, reloaded.CustomDocumentProperties["ReviewedAt"].Date);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBoldItalicRunsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("plain ");
                    paragraph.AddText("bold ").SetBold();
                    paragraph.AddText("italic").SetItalic();

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] runs = reloaded.Paragraphs.ToArray();
                Assert.Equal(3, runs.Length);
                Assert.Equal("plain ", runs[0].Text);
                Assert.False(runs[0].Bold);
                Assert.False(runs[0].Italic);
                Assert.Equal("bold ", runs[1].Text);
                Assert.True(runs[1].Bold);
                Assert.False(runs[1].Italic);
                Assert.Equal("italic", runs[2].Text);
                Assert.False(runs[2].Bold);
                Assert.True(runs[2].Italic);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocUnderlineSizeAndColorRunsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("plain ");
                    paragraph.AddText("under ").SetUnderline(UnderlineValues.Single);
                    paragraph.AddText("sized ").SetFontSize(14);
                    paragraph.AddText("color").SetColorHex("336699");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] runs = reloaded.Paragraphs.ToArray();
                Assert.Equal(4, runs.Length);
                Assert.Equal("plain ", runs[0].Text);
                Assert.Null(runs[0].Underline);
                Assert.Equal("under ", runs[1].Text);
                Assert.Equal(UnderlineValues.Single, runs[1].Underline);
                Assert.Equal("sized ", runs[2].Text);
                Assert.Equal(14, runs[2].FontSize);
                Assert.Equal("color", runs[3].Text);
                Assert.Equal("336699", runs[3].ColorHex);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocFontFamilyRunsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("plain ");
                    paragraph.AddText("font").SetFontFamily("Courier New");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] runs = reloaded.Paragraphs.ToArray();
                Assert.Equal(2, runs.Length);
                Assert.Equal("plain ", runs[0].Text);
                Assert.Null(runs[0].FontFamily);
                Assert.Equal("font", runs[1].Text);
                Assert.Equal("Courier New", runs[1].FontFamily);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphAlignmentAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("left");
                    document.AddParagraph("center").ParagraphAlignment = JustificationValues.Center;
                    document.AddParagraph("right").ParagraphAlignment = JustificationValues.Right;

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] paragraphs = reloaded.Paragraphs.ToArray();
                Assert.Equal(3, paragraphs.Length);
                Assert.Equal("left", paragraphs[0].Text);
                Assert.Null(paragraphs[0].ParagraphAlignment);
                Assert.Equal("center", paragraphs[1].Text);
                Assert.Equal(JustificationValues.Center, paragraphs[1].ParagraphAlignment);
                Assert.Equal("right", paragraphs[2].Text);
                Assert.Equal(JustificationValues.Right, paragraphs[2].ParagraphAlignment);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksUnsupportedRunFormattingBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                document.AddParagraph("Formatted").SetHighlight(HighlightColorValues.Yellow);

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("highlight", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static class LegacyDocTestBuilder {
            internal static byte[] CreateSimpleDoc(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithDocumentProperties(DateTime created, DateTime modified, params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);
                byte[] summaryInformation = CreateSummaryInformationPropertySet(created, modified);
                byte[] documentSummaryInformation = CreateDocumentSummaryInformationPropertySet();

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                    WriteStream(root, "\u0005SummaryInformation", summaryInformation);
                    WriteStream(root, "\u0005DocumentSummaryInformation", documentSummaryInformation);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithUnsupportedFeatureStorage(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                    root.CreateStorage("_VBA_PROJECT_CUR");
                    root.CreateStorage("ObjectPool");
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithDirectCharacterFormatting() {
                const string text = "plain bold italic\r";
                const int textOffset = 0x200;
                const int chpxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithDirectCharacterFormatting(text, textOffset, chpxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithCharacterBinTable(text.Length, textOffset, chpxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithExtendedDirectCharacterFormatting() {
                const string text = "plain under sized red direct\r";
                const int textOffset = 0x200;
                const int chpxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithExtendedCharacterFormatting(text, textOffset, chpxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithCharacterBinTable(text.Length, textOffset, chpxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithFontFamilyFormatting() {
                const string text = "plain font\r";
                const int textOffset = 0x200;
                const int chpxFkpOffset = 0x400;
                byte[] fontTable = CreateFontTable("Courier New");
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithFontFamilyFormatting(text, textOffset, chpxFkpOffset, fontTable.Length);
                byte[] tableStream = CreateUnicodeTableStreamWithCharacterBinTableAndFontTable(text.Length, textOffset, chpxFkpOffset / 512, fontTable);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithParagraphAlignment() {
                const string text = "left\rcenter\rright\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithParagraphAlignment(text, textOffset, papxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateCompoundWithoutWordDocumentStream() {
                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "NotWordDocument", new byte[] { 1, 2, 3 });
                }

                return package.ToArray();
            }

            private static byte[] CreateWordDocumentStream(string text) {
                const int fibLength = 0x1AA;
                const int textOffset = 0x200;
                byte[] textBytes = EncodeWindows1252(text);
                var stream = new byte[textOffset + textBytes.Length];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);
                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithDirectCharacterFormatting(string text, int textOffset, int chpxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(chpxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xFA, 21);
                WriteInt32(stream, 0xFE, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int boldStart = textOffset + ("plain ".Length * 2);
                int italicStart = boldStart + ("bold ".Length * 2);
                int paragraphMarkStart = italicStart + ("italic".Length * 2);
                int end = paragraphMarkStart + 2;
                WriteChpxFkp(
                    stream,
                    chpxFkpOffset,
                    new[] { textOffset, boldStart, italicStart, paragraphMarkStart, end },
                    boldRunIndex: 1,
                    italicRunIndex: 2);

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithExtendedCharacterFormatting(string text, int textOffset, int chpxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(chpxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xFA, 21);
                WriteInt32(stream, 0xFE, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int underStart = textOffset + ("plain ".Length * 2);
                int sizedStart = underStart + ("under ".Length * 2);
                int redStart = sizedStart + ("sized ".Length * 2);
                int directStart = redStart + ("red ".Length * 2);
                int paragraphMarkStart = directStart + ("direct".Length * 2);
                int end = paragraphMarkStart + 2;
                WriteChpxFkp(
                    stream,
                    chpxFkpOffset,
                    new[] { textOffset, underStart, sizedStart, redStart, directStart, paragraphMarkStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateSingleSprmChpx(0x2A3E, 1),
                        [2] = CreateSingleSprmChpx(0x4A43, 28, 0),
                        [3] = CreateSingleSprmChpx(0x2A42, 6),
                        [4] = CreateSingleSprmChpx(0x6870, 0x33, 0x66, 0x99, 0)
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithFontFamilyFormatting(string text, int textOffset, int chpxFkpOffset, int fontTableLength) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(chpxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xFA, 21);
                WriteInt32(stream, 0xFE, 12);
                WriteInt32(stream, 0x112, 33);
                WriteInt32(stream, 0x116, fontTableLength);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int fontStart = textOffset + ("plain ".Length * 2);
                int paragraphMarkStart = fontStart + ("font".Length * 2);
                int end = paragraphMarkStart + 2;
                WriteChpxFkp(
                    stream,
                    chpxFkpOffset,
                    new[] { textOffset, fontStart, paragraphMarkStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateSingleSprmChpx(0x4A4F, 0, 0)
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithParagraphAlignment(string text, int textOffset, int papxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int centerStart = textOffset + ("left\r".Length * 2);
                int rightStart = centerStart + ("center\r".Length * 2);
                int end = rightStart + ("right\r".Length * 2);
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, centerStart, rightStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateParagraphAlignmentPapx(1),
                        [2] = CreateParagraphAlignmentPapx(2)
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateTableStream(int characterCount) {
                const int textOffset = 0x200;
                var table = new byte[21];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, 0x40000000U | ((uint)textOffset * 2U));
                WriteUInt16(table, 19, 0);
                return table;
            }

            private static byte[] CreateUnicodeTableStreamWithParagraphBinTable(int characterCount, int textOffset, int papxFkpPageNumber) {
                var table = new byte[33];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, unchecked((uint)textOffset));
                WriteUInt16(table, 19, 0);

                int papxPlcOffset = 21;
                WriteInt32(table, papxPlcOffset, textOffset);
                WriteInt32(table, papxPlcOffset + 4, textOffset + (characterCount * 2));
                WriteInt32(table, papxPlcOffset + 8, papxFkpPageNumber);
                return table;
            }

            private static byte[] CreateUnicodeTableStreamWithCharacterBinTableAndFontTable(int characterCount, int textOffset, int chpxFkpPageNumber, byte[] fontTable) {
                byte[] table = CreateUnicodeTableStreamWithCharacterBinTable(characterCount, textOffset, chpxFkpPageNumber);
                Array.Resize(ref table, table.Length + fontTable.Length);
                Buffer.BlockCopy(fontTable, 0, table, 33, fontTable.Length);
                return table;
            }

            private static byte[] CreateFontTable(params string[] fontFamilies) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)fontFamilies.Length));
                WriteUInt16(stream, 0);
                foreach (string fontFamily in fontFamilies) {
                    byte[] ffn = CreateFfn(fontFamily);
                    stream.WriteByte(checked((byte)ffn.Length));
                    stream.Write(ffn, 0, ffn.Length);
                }

                return stream.ToArray();
            }

            private static byte[] CreateFfn(string fontFamily) {
                byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(fontFamily + '\0');
                var ffn = new byte[39 + nameBytes.Length];
                ffn[1] = 0x90;
                ffn[2] = 0x01;
                Buffer.BlockCopy(nameBytes, 0, ffn, 39, nameBytes.Length);
                return ffn;
            }

            private static byte[] CreateUnicodeTableStreamWithCharacterBinTable(int characterCount, int textOffset, int chpxFkpPageNumber) {
                var table = new byte[33];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, unchecked((uint)textOffset));
                WriteUInt16(table, 19, 0);

                int chpxPlcOffset = 21;
                WriteInt32(table, chpxPlcOffset, textOffset);
                WriteInt32(table, chpxPlcOffset + 4, textOffset + (characterCount * 2));
                WriteInt32(table, chpxPlcOffset + 8, chpxFkpPageNumber);
                return table;
            }

            private static void WriteChpxFkp(byte[] stream, int fkpOffset, int[] fileCharacterPositions, int boldRunIndex, int italicRunIndex) {
                const int boldChpxOffset = 0xF0;
                const int italicChpxOffset = 0xF8;
                int runCount = fileCharacterPositions.Length - 1;
                for (int i = 0; i < fileCharacterPositions.Length; i++) {
                    WriteInt32(stream, fkpOffset + (i * 4), fileCharacterPositions[i]);
                }

                int rgbOffset = fkpOffset + (fileCharacterPositions.Length * 4);
                for (int i = 0; i < runCount; i++) {
                    if (i == boldRunIndex) {
                        stream[rgbOffset + i] = boldChpxOffset / 2;
                    } else if (i == italicRunIndex) {
                        stream[rgbOffset + i] = italicChpxOffset / 2;
                    }
                }

                WriteSingleToggleChpx(stream, fkpOffset + boldChpxOffset, 0x0835);
                WriteSingleToggleChpx(stream, fkpOffset + italicChpxOffset, 0x0836);
                stream[fkpOffset + 511] = checked((byte)runCount);
            }

            private static void WriteChpxFkp(byte[] stream, int fkpOffset, int[] fileCharacterPositions, IReadOnlyDictionary<int, byte[]> chpxByRunIndex) {
                int runCount = fileCharacterPositions.Length - 1;
                for (int i = 0; i < fileCharacterPositions.Length; i++) {
                    WriteInt32(stream, fkpOffset + (i * 4), fileCharacterPositions[i]);
                }

                int rgbOffset = fkpOffset + (fileCharacterPositions.Length * 4);
                int chpxOffset = 0xE0;
                for (int i = 0; i < runCount; i++) {
                    if (!chpxByRunIndex.TryGetValue(i, out byte[]? chpx)) {
                        continue;
                    }

                    chpxOffset = AlignToEven(chpxOffset);
                    stream[rgbOffset + i] = checked((byte)(chpxOffset / 2));
                    Buffer.BlockCopy(chpx, 0, stream, fkpOffset + chpxOffset, chpx.Length);
                    chpxOffset += chpx.Length;
                }

                stream[fkpOffset + 511] = checked((byte)runCount);
            }

            private static void WritePapxFkp(byte[] stream, int fkpOffset, int[] fileParagraphPositions, IReadOnlyDictionary<int, byte[]> papxByParagraphIndex) {
                const int bxLength = 13;
                int paragraphCount = fileParagraphPositions.Length - 1;
                for (int i = 0; i < fileParagraphPositions.Length; i++) {
                    WriteInt32(stream, fkpOffset + (i * 4), fileParagraphPositions[i]);
                }

                int rgbxOffset = fkpOffset + (fileParagraphPositions.Length * 4);
                int papxOffset = 0x1E0;
                for (int i = 0; i < paragraphCount; i++) {
                    if (!papxByParagraphIndex.TryGetValue(i, out byte[]? papx)) {
                        continue;
                    }

                    papxOffset = AlignToEven(papxOffset);
                    stream[rgbxOffset + (i * bxLength)] = checked((byte)(papxOffset / 2));
                    Buffer.BlockCopy(papx, 0, stream, fkpOffset + papxOffset, papx.Length);
                    papxOffset += papx.Length;
                }

                stream[fkpOffset + 511] = checked((byte)paragraphCount);
            }

            private static void WriteSingleToggleChpx(byte[] stream, int offset, ushort sprm) {
                stream[offset] = 3;
                WriteUInt16(stream, offset + 1, sprm);
                stream[offset + 3] = 1;
            }

            private static byte[] CreateSingleSprmChpx(ushort sprm, params byte[] operand) {
                var chpx = new byte[3 + operand.Length];
                chpx[0] = checked((byte)(2 + operand.Length));
                WriteUInt16(chpx, 1, sprm);
                Buffer.BlockCopy(operand, 0, chpx, 3, operand.Length);
                return chpx;
            }

            private static byte[] CreateParagraphAlignmentPapx(byte alignment) {
                return new byte[] {
                    0,
                    3,
                    0,
                    0,
                    0x61,
                    0x24,
                    alignment,
                    0
                };
            }

            private static int AlignToEven(int value) {
                return value % 2 == 0 ? value : value + 1;
            }

            private static byte[] CreateSummaryInformationPropertySet(DateTime created, DateTime modified) {
                var properties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.String(2, "Legacy DOC Metadata Title"),
                    OleTestProperty.String(3, "Legacy DOC metadata subject"),
                    OleTestProperty.String(4, "OfficeIMO Legacy Import"),
                    OleTestProperty.String(5, "doc, metadata, officeimo"),
                    OleTestProperty.String(6, "OLE SummaryInformation comments"),
                    OleTestProperty.FileTime(12, created),
                    OleTestProperty.FileTime(13, modified)
                };

                return CreateOlePropertySet(CreateOlePropertySection(properties));
            }

            private static byte[] CreateDocumentSummaryInformationPropertySet() {
                var documentProperties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.String(2, "Legacy Category"),
                    OleTestProperty.String(14, "Document Manager"),
                    OleTestProperty.String(15, "EvotecIT")
                };
                var customProperties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.Dictionary(0, new Dictionary<uint, string> {
                        [2] = "ReleaseStatus",
                        [3] = "Reviewed",
                        [4] = "Ticket"
                    }),
                    OleTestProperty.String(2, "Ready"),
                    OleTestProperty.Boolean(3, true),
                    OleTestProperty.Int32(4, 2003)
                };

                return CreateOlePropertySet(CreateOlePropertySection(documentProperties), CreateOlePropertySection(customProperties));
            }

            private static byte[] CreateOlePropertySet(params byte[][] sections) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0xfffe);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, checked((uint)sections.Length));

                int sectionOffset = 28 + sections.Length * 20;
                foreach (byte[] section in sections) {
                    stream.Write(new byte[16], 0, 16);
                    WriteUInt32(stream, checked((uint)sectionOffset));
                    sectionOffset += section.Length;
                }

                foreach (byte[] section in sections) {
                    stream.Write(section, 0, section.Length);
                }

                return stream.ToArray();
            }

            private static byte[] CreateOlePropertySection(IReadOnlyList<OleTestProperty> properties) {
                using var values = new MemoryStream();
                var offsets = new List<uint>(properties.Count);
                foreach (OleTestProperty property in properties) {
                    offsets.Add(checked((uint)(8 + properties.Count * 8 + values.Length)));
                    values.Write(property.ValueBytes, 0, property.ValueBytes.Length);
                    PadToInt32(values);
                }

                using var stream = new MemoryStream();
                WriteUInt32(stream, checked((uint)(8 + properties.Count * 8 + values.Length)));
                WriteUInt32(stream, checked((uint)properties.Count));
                for (int i = 0; i < properties.Count; i++) {
                    WriteUInt32(stream, properties[i].PropertyId);
                    WriteUInt32(stream, offsets[i]);
                }

                byte[] valueBytes = values.ToArray();
                stream.Write(valueBytes, 0, valueBytes.Length);
                return stream.ToArray();
            }

            private static void WriteStream(RootStorage root, string name, byte[] bytes) {
                using CfbStream stream = root.CreateStream(name);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static byte[] EncodeWindows1252(string text) {
                var bytes = new byte[text.Length];
                for (int i = 0; i < text.Length; i++) {
                    char character = text[i];
                    bytes[i] = character <= 0x7F || (character >= 0xA0 && character <= 0xFF)
                        ? (byte)character
                        : (byte)'?';
                }

                return bytes;
            }

            private static void PadToInt32(Stream stream) {
                while (stream.Position % 4 != 0) {
                    stream.WriteByte(0);
                }
            }

            private static void WriteUInt16(Stream stream, ushort value) {
                stream.WriteByte((byte)(value & 0xff));
                stream.WriteByte((byte)((value >> 8) & 0xff));
            }

            private static void WriteUInt32(Stream stream, uint value) {
                stream.WriteByte((byte)(value & 0xff));
                stream.WriteByte((byte)((value >> 8) & 0xff));
                stream.WriteByte((byte)((value >> 16) & 0xff));
                stream.WriteByte((byte)((value >> 24) & 0xff));
            }

            private static void WriteUInt64(Stream stream, ulong value) {
                WriteUInt32(stream, unchecked((uint)(value & 0xffffffffUL)));
                WriteUInt32(stream, unchecked((uint)(value >> 32)));
            }

            private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
            }

            private static void WriteInt32(byte[] bytes, int offset, int value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }

            private static void WriteUInt32(byte[] bytes, int offset, uint value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }

            private readonly struct OleTestProperty {
                private OleTestProperty(uint propertyId, byte[] valueBytes) {
                    PropertyId = propertyId;
                    ValueBytes = valueBytes;
                }

                internal uint PropertyId { get; }

                internal byte[] ValueBytes { get; }

                internal static OleTestProperty Int16(uint id, short value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0002);
                    WriteUInt16(stream, 0);
                    WriteUInt16(stream, unchecked((ushort)value));
                    WriteUInt16(stream, 0);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Int32(uint id, int value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0003);
                    WriteUInt16(stream, 0);
                    WriteUInt32(stream, unchecked((uint)value));
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Boolean(uint id, bool value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x000b);
                    WriteUInt16(stream, 0);
                    WriteUInt16(stream, value ? (ushort)0xffff : (ushort)0);
                    WriteUInt16(stream, 0);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty FileTime(uint id, DateTime value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0040);
                    WriteUInt16(stream, 0);
                    WriteUInt64(stream, unchecked((ulong)value.ToUniversalTime().ToFileTimeUtc()));
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty String(uint id, string value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x001f);
                    WriteUInt16(stream, 0);
                    WriteUInt32(stream, checked((uint)(value.Length + 1)));
                    byte[] bytes = System.Text.Encoding.Unicode.GetBytes(value + '\0');
                    stream.Write(bytes, 0, bytes.Length);
                    PadToInt32(stream);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Dictionary(uint id, IReadOnlyDictionary<uint, string> names) {
                    using var stream = new MemoryStream();
                    WriteUInt32(stream, checked((uint)names.Count));
                    foreach (KeyValuePair<uint, string> name in names.OrderBy(entry => entry.Key)) {
                        WriteUInt32(stream, name.Key);
                        WriteUInt32(stream, checked((uint)(name.Value.Length + 1)));
                        byte[] bytes = System.Text.Encoding.Unicode.GetBytes(name.Value + '\0');
                        stream.Write(bytes, 0, bytes.Length);
                        PadToInt32(stream);
                    }

                    return new OleTestProperty(id, stream.ToArray());
                }
            }
        }

        private static void AssertSameInstant(DateTime expected, DateTime? actual) {
            Assert.NotNull(actual);
            Assert.Equal(expected.ToUniversalTime(), actual.Value.ToUniversalTime());
        }

        private static string NormalizeLegacyDocBaselineText(string text) {
            return text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd() + "\n";
        }

        private static string GetWordTestsProjectRoot() {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Tests.csproj"))) {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return AppContext.BaseDirectory;
        }

        private static void DeleteIfExists(string path) {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }
}
