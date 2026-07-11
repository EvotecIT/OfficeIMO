using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_Convert_DocxToDocAndBack_RoundTripsSupportedContent() {
            string docxPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string docPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string roundTripPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");

            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Conversion first paragraph");
                document.AddParagraph("Conversion second paragraph");
                document.Save(docxPath);
            }

            WordDocumentConversionResult toDoc = WordDocument.Convert(docxPath, docPath);

            Assert.True(toDoc.OutputCreated);
            Assert.Equal(WordFileFormat.Docx, toDoc.SourceFormat);
            Assert.Equal(WordFileFormat.Doc, toDoc.DestinationFormat);
            Assert.False(toDoc.HasDataLoss);

            AssertOleCompoundFile(docPath);
            using (WordDocument legacy = WordDocument.Load(docPath)) {
                Assert.True(legacy.SourceFormat == WordFileFormat.Doc);
                Assert.Empty(legacy.LegacyDocUnsupportedFeatures);
                Assert.Contains(legacy.Paragraphs, paragraph => paragraph.Text == "Conversion first paragraph");
                Assert.Contains(legacy.Paragraphs, paragraph => paragraph.Text == "Conversion second paragraph");
            }

            WordDocumentConversionResult toDocx = WordDocument.Convert(docPath, roundTripPath);

            Assert.Equal(WordFileFormat.Doc, toDocx.SourceFormat);
            Assert.Equal(WordFileFormat.Docx, toDocx.DestinationFormat);

            using WordDocument roundTrip = WordDocument.Load(roundTripPath);
            Assert.False(roundTrip.SourceFormat == WordFileFormat.Doc);
            Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.Text == "Conversion first paragraph");
            Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.Text == "Conversion second paragraph");
        }

        [Fact]
        public void LegacyDoc_Convert_BlocksUnsupportedLegacyContentUnlessLossIsAllowed() {
            string docPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Preserve-only body"));

            WordDocumentConversionException exception = Assert.Throws<WordDocumentConversionException>(() => WordDocument.Convert(docPath, blockedPath));

            Assert.Equal(WordDocumentConversionFailureReason.DataLossBlocked, exception.Reason);
            Assert.True(exception.Result.HasDataLoss);
            Assert.Contains(exception.Result.Diagnostics, diagnostic => diagnostic.Category == WordConversionDiagnosticCategory.DataLoss);
            Assert.False(File.Exists(blockedPath));

            WordDocumentConversionResult result = WordDocument.Convert(docPath, allowedPath, new WordDocumentConversionOptions {
                LossPolicy = WordConversionLossPolicy.Allow
            });

            Assert.True(result.HasDataLoss);
            Assert.True(result.OutputCreated);

            using WordDocument converted = WordDocument.Load(allowedPath);
            Assert.False(converted.SourceFormat == WordFileFormat.Doc);
            Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text == "Preserve-only body");
        }

        [Fact]
        public void LegacyDoc_Convert_ContentDetectionAppliesImportOptionsDespiteMisleadingExtension() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            File.WriteAllBytes(sourcePath, LegacyDocTestBuilder.CreateSimpleDoc("Physical DOC"));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                WordDocument.Convert(sourcePath, destinationPath, new WordDocumentConversionOptions {
                    LossPolicy = WordConversionLossPolicy.Allow,
                    LegacyDocImportOptions = new LegacyDocImportOptions { MaxInputBytes = 1 }
                }));

            Assert.Contains("configured limit of 1 byte", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void LegacyDoc_Convert_CannotSuppressDiscoveryToBypassLossPolicy() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            File.WriteAllBytes(sourcePath, LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Safety gate"));

            WordDocumentConversionException exception = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(sourcePath, destinationPath, new WordDocumentConversionOptions {
                    LegacyDocImportOptions = new LegacyDocImportOptions { ReportUnsupportedContent = false }
                }));

            Assert.Equal(WordDocumentConversionFailureReason.DataLossBlocked, exception.Reason);
            Assert.True(exception.Result.HasDataLoss);
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void LegacyDoc_Convert_DefaultConflictPolicyPreservesExistingDestination() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            byte[] existing = { 1, 2, 3, 4 };
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Conflict policy");
                document.Save(sourcePath);
            }
            File.WriteAllBytes(destinationPath, existing);

            WordDocumentConversionException exception = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(sourcePath, destinationPath));

            Assert.Equal(WordDocumentConversionFailureReason.DestinationExists, exception.Reason);
            Assert.Equal(existing, File.ReadAllBytes(destinationPath));

            WordDocumentConversionResult replaced = WordDocument.Convert(sourcePath, destinationPath, new WordDocumentConversionOptions {
                FileConflictPolicy = WordConversionFileConflictPolicy.Replace
            });
            Assert.True(replaced.ReplacedExistingFile);
            AssertOleCompoundFile(destinationPath);
        }

        [Fact]
        public void LegacyDoc_Convert_ReplacePreservesReadOnlyDestination() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            byte[] originalBytes = { 1, 2, 3, 4 };
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Read-only conversion target");
                document.Save(sourcePath);
            }
            File.WriteAllBytes(destinationPath, originalBytes);
            var destination = new FileInfo(destinationPath) { IsReadOnly = true };

            try {
                Assert.Throws<IOException>(() => WordDocument.Convert(sourcePath, destinationPath, new WordDocumentConversionOptions {
                    FileConflictPolicy = WordConversionFileConflictPolicy.Replace
                }));

                Assert.Equal(originalBytes, File.ReadAllBytes(destinationPath));
            } finally {
                destination.IsReadOnly = false;
            }
        }

        [Fact]
        public void LegacyDoc_Convert_RejectsSamePhysicalFormat() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Already DOCX");
                document.Save(sourcePath);
            }

            WordDocumentConversionException exception = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(sourcePath, destinationPath));

            Assert.Equal(WordDocumentConversionFailureReason.SameFormat, exception.Reason);
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void LegacyDoc_Convert_BlocksPreservedLegacyMetadataUnlessLossIsAllowed() {
            string docPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            const ushort pictureFlag = 0x0008;
            File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDocWithFibFlags(pictureFlag, "Picture body"));

            WordDocumentConversionException exception = Assert.Throws<WordDocumentConversionException>(() => WordDocument.Convert(docPath, blockedPath));

            Assert.Equal(WordDocumentConversionFailureReason.DataLossBlocked, exception.Reason);
            Assert.True(exception.Result.HasDataLoss);
            Assert.False(File.Exists(blockedPath));

            WordDocument.Convert(docPath, allowedPath, new WordDocumentConversionOptions {
                LossPolicy = WordConversionLossPolicy.Allow
            });

            using WordDocument converted = WordDocument.Load(allowedPath);
            Assert.False(converted.SourceFormat == WordFileFormat.Doc);
            Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text == "Picture body");
        }

        [Fact]
        public void LegacyDoc_NormalSave_BlocksKnownImportLossUnlessExplicitlyAllowed() {
            string docPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Normal save loss gate"));

            using WordDocument document = WordDocument.Load(docPath);

            Assert.Throws<NotSupportedException>(() => document.Save(blockedPath));
            Assert.False(File.Exists(blockedPath));

            document.Save(allowedPath, new WordSaveOptions {
                LossPolicy = WordConversionLossPolicy.Allow
            });

            using WordDocument saved = WordDocument.Load(allowedPath);
            Assert.Contains(saved.Paragraphs, paragraph => paragraph.Text == "Normal save loss gate");
        }

        private static void AssertOleCompoundFile(string path) {
            byte[] bytes = File.ReadAllBytes(path);
            Assert.True(bytes.Length > 8);
            Assert.Equal(0xd0, bytes[0]);
            Assert.Equal(0xcf, bytes[1]);
            Assert.Equal(0x11, bytes[2]);
            Assert.Equal(0xe0, bytes[3]);
        }
    }
}
