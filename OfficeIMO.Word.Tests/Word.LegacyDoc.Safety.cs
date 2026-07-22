using OfficeIMO.Drawing.Internal;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_ToDoc_ProducesDocBytes() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Native DOC byte output");

            byte[] bytes = document.ToBytes(OfficeIMO.Word.WordFileFormat.Doc);

            AssertOleCompoundBytes(bytes);
            using WordDocument reloaded = WordDocument.Load(new MemoryStream(bytes));
            Assert.Equal(WordFileFormat.Doc, reloaded.SourceFormat);
            Assert.Contains(reloaded.Paragraphs, paragraph => paragraph.Text == "Native DOC byte output");
        }

        [Fact]
        public async Task FormatApi_ToDocxAndLoadAsyncStream_RoundTrips() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Async DOCX stream");
            byte[] bytes = document.ToBytes();

            Assert.Equal(0x50, bytes[0]);
            Assert.Equal(0x4b, bytes[1]);
            using var stream = new MemoryStream(bytes);
            using WordDocument loaded = await WordDocument.LoadAsync(stream);

            Assert.Equal(WordFileFormat.Docx, loaded.SourceFormat);
            Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text == "Async DOCX stream");
        }

        [Fact]
        public async Task Load_StreamAndAsyncStreamEnforceCompleteInputLimit() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Bounded DOCX stream");
            byte[] bytes = document.ToBytes();

            using var syncStream = new MemoryStream(bytes);
            Assert.Throws<InvalidDataException>(() => WordDocument.Load(
                syncStream,
                new WordLoadOptions { MaxInputBytes = bytes.Length - 1L }));

            using var asyncStream = new MemoryStream(bytes);
            await Assert.ThrowsAsync<InvalidDataException>(() => WordDocument.LoadAsync(
                asyncStream,
                new WordLoadOptions { MaxInputBytes = bytes.Length - 1L }));
        }

        [Fact]
        public void LegacyDoc_LoadResult_CachesCompactAndAdvancedReports() {
            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(
                new MemoryStream(LegacyDocTestBuilder.CreateSimpleDoc("Summary paragraph")));

            Assert.Equal(1, result.Summary.ParagraphCount);
            Assert.False(result.Summary.HasImportErrors);
            Assert.Same(result.Summary, result.Summary);
            Assert.Same(result.ImportReport, result.CreateAdvancedImportReport());
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDoc_PathAssociatesSourceAndSavePersistsDoc() {
            string path = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            File.WriteAllBytes(path, LegacyDocTestBuilder.CreateSimpleDoc("Original paragraph"));

            using (WordDocument document = WordDocument.LoadLegacyDoc(path)) {
                Assert.Equal(path, document.FilePath);
                Assert.Equal(path, document.SourcePath);
                Assert.Equal(WordFileFormat.Doc, document.SourceFormat);
                document.AddParagraph("Saved paragraph");
                document.Save();
            }

            AssertOleCompoundFile(path);
            using WordDocument reloaded = WordDocument.LoadLegacyDoc(path);
            Assert.Contains(reloaded.Paragraphs, paragraph => paragraph.Text == "Original paragraph");
            Assert.Contains(reloaded.Paragraphs, paragraph => paragraph.Text == "Saved paragraph");
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDoc_StreamSaveWithoutDestinationThrows() {
            using WordDocument document = WordDocument.LoadLegacyDoc(new MemoryStream(LegacyDocTestBuilder.CreateSimpleDoc("Stream paragraph")));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => document.Save());

            Assert.Contains("not associated with a file path", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Load_WhenInputIsLegacyExcel_ReportsFormatMismatch() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "LegacyXlsCorpus",
                "apache-poi-testdata",
                "SimpleWithComments.xls");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => WordDocument.Load(path));

            Assert.Contains("legacy Excel workbook", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Load_WhenInputIsLegacyPowerPoint_ReportsFormatMismatch() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "LegacyPptCorpus",
                "BasicPowerPoint.ppt");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => WordDocument.Load(path));

            Assert.Contains("legacy PowerPoint presentation", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void FormatApi_UsesCanonicalWordNamingWithoutLegacyAliases() {
            Type documentType = typeof(WordDocument);
            Type importOptionsType = typeof(LegacyDocImportOptions);

            Assert.NotNull(documentType.GetMethod(nameof(WordDocument.ToBytes),
                new[] { typeof(WordFileFormat), typeof(WordSaveOptions) }));
            Assert.NotNull(documentType.GetMethod(nameof(WordDocument.ToStream),
                new[] { typeof(WordFileFormat), typeof(WordSaveOptions) }));
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name is "ToDocx" or "ToDoc" or "ToDocxStream" or "ToDocStream");
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name is "Save" or "SaveAsync" &&
                method.GetParameters().Any(parameter => parameter.ParameterType == typeof(bool)));
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name == "Open");
            Assert.Contains(documentType.GetMethods(), method => method.Name == nameof(WordDocument.OpenInApplication));
            Assert.Contains(documentType.GetMethods(), method => method.Name == nameof(WordDocument.SaveCopy));
            Assert.Null(documentType.GetMethod("SaveAsByteArray"));
            Assert.Null(documentType.GetMethod("SaveAsMemoryStream"));
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name == "SaveAs" && method.ReturnType == typeof(WordDocument));
            Assert.Null(documentType.GetProperty("WasLoadedFromLegacyDoc"));
            Assert.Null(importOptionsType.GetProperty("MaxWordDocumentStreamBytes"));
            Assert.Null(importOptionsType.GetProperty("ReportUnsupportedFeatures"));
            Assert.NotNull(importOptionsType.GetProperty(nameof(LegacyDocImportOptions.MaxInputBytes)));
            Assert.NotNull(importOptionsType.GetProperty(nameof(LegacyDocImportOptions.MaxDecodedImageBytes)));
            Assert.NotNull(importOptionsType.GetProperty(nameof(LegacyDocImportOptions.ReportUnsupportedContent)));
        }

        [Fact]
        public void LegacyDoc_RejectsNonPositiveDecodedImageBudget() {
            byte[] bytes = LegacyDocTestBuilder.CreateSimpleDoc("Image budget");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                OfficeIMO.Word.LegacyDoc.Model.LegacyDocDocument.Load(bytes,
                    new LegacyDocImportOptions { MaxDecodedImageBytes = 0 }));
        }

        [Fact]
        public void LegacyDoc_AllPublicLoadShapesEnforceTheCompleteInputBudget() {
            byte[] bytes = LegacyDocTestBuilder.CreateSimpleDoc("Bounded legacy input");
            var options = new LegacyDocImportOptions { MaxInputBytes = bytes.Length - 1 };
            string path = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            File.WriteAllBytes(path, bytes);

            Assert.Throws<InvalidDataException>(() =>
                OfficeIMO.Word.LegacyDoc.Model.LegacyDocDocument.Load(bytes, options));
            Assert.Throws<InvalidDataException>(() =>
                OfficeIMO.Word.LegacyDoc.Model.LegacyDocDocument.Load(new MemoryStream(bytes), options));
            Assert.Throws<InvalidDataException>(() =>
                OfficeIMO.Word.LegacyDoc.Model.LegacyDocDocument.Load(path, options));
        }

        private static void AssertOleCompoundBytes(byte[] bytes) {
            Assert.True(bytes.Length > 8);
            Assert.Equal(0xd0, bytes[0]);
            Assert.Equal(0xcf, bytes[1]);
            Assert.Equal(0x11, bytes[2]);
            Assert.Equal(0xe0, bytes[3]);
        }
    }
}
