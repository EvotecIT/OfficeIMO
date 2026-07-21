using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
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

            Assert.Equal(docPath, toDoc.RequireNoLoss());
            Assert.Equal(WordFileFormat.Docx, toDoc.Report.SourceFormat);
            Assert.Equal(WordFileFormat.Doc, toDoc.Report.DestinationFormat);
            Assert.False(toDoc.HasLoss);

            AssertOleCompoundFile(docPath);
            using (WordDocument legacy = WordDocument.Load(docPath)) {
                Assert.True(legacy.SourceFormat == WordFileFormat.Doc);
                Assert.Empty(legacy.LegacyDocUnsupportedFeatures);
                Assert.Contains(legacy.Paragraphs, paragraph => paragraph.Text == "Conversion first paragraph");
                Assert.Contains(legacy.Paragraphs, paragraph => paragraph.Text == "Conversion second paragraph");
            }

            WordDocumentConversionResult toDocx = WordDocument.Convert(docPath, roundTripPath);

            Assert.Equal(WordFileFormat.Doc, toDocx.Report.SourceFormat);
            Assert.Equal(WordFileFormat.Docx, toDocx.Report.DestinationFormat);

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
            Assert.True(exception.Result.HasLoss);
            Assert.Contains(exception.Result.Report.Diagnostics, diagnostic => diagnostic.Category == WordConversionDiagnosticCategory.DataLoss);
            Assert.False(File.Exists(blockedPath));

            WordDocumentConversionResult result = WordDocument.Convert(docPath, allowedPath, new WordDocumentConversionOptions {
                LossPolicy = WordConversionLossPolicy.Allow
            });

            Assert.True(result.HasLoss);
            Assert.Equal(allowedPath, result.RequireValue());
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());

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
            Assert.True(exception.Result.HasLoss);
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
            Assert.True(replaced.Report.ReplacedExistingFile);
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
        public void Convert_ModernDocumentToTemplate_UsesConcreteFormatDescriptors() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".dotx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Reusable template content");
                document.Save();
            }

            WordDocumentConversionResult result = WordDocument.Convert(sourcePath, destinationPath);

            Assert.Equal("Word.Docx", result.Report.SourceFormatDescriptor.Id);
            Assert.Equal("Word.Dotx", result.Report.DestinationFormatDescriptor.Id);
            Assert.Equal(OfficeDocumentKind.Template, result.Report.DestinationFormatDescriptor.DocumentKind);
            Assert.Equal(destinationPath, result.RequireNoLoss());
            using WordprocessingDocument package = WordprocessingDocument.Open(destinationPath, false);
            Assert.Equal(WordprocessingDocumentType.Template, package.DocumentType);
        }

        [Fact]
        public void Convert_ModernDocumentToBinaryTemplate_WritesTemplateFibAndRoundTripsKind() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string templatePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".dot");
            string roundTripPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Binary template content");
                document.Save();
            }

            WordDocumentConversionResult template = WordDocument.Convert(sourcePath, templatePath);

            Assert.Equal("Word.Dot", template.Report.DestinationFormatDescriptor.Id);
            Assert.Equal(OfficeDocumentKind.Template, template.Report.DestinationFormatDescriptor.DocumentKind);
            byte[] wordDocumentStream = ReadCompoundStream(File.ReadAllBytes(templatePath), "WordDocument");
            ushort fibFlags = BitConverter.ToUInt16(wordDocumentStream, 0x0A);
            Assert.NotEqual(0, fibFlags & 0x0001);

            WordDocumentConversionResult roundTrip = WordDocument.Convert(templatePath, roundTripPath);
            Assert.Equal("Word.Dot", roundTrip.Report.SourceFormatDescriptor.Id);
            using WordDocument converted = WordDocument.Load(roundTripPath);
            Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text == "Binary template content");
        }

        [Fact]
        public void AnalyzeConversion_ReturnsSamePolicySurfaceWithoutCreatingOutput() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".dot");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Preflight only");
                document.Save();
            }

            WordDocumentConversionReport report = WordDocument.AnalyzeConversion(sourcePath, destinationPath);

            Assert.Equal("Word.Docx", report.SourceFormatDescriptor.Id);
            Assert.Equal("Word.Dot", report.DestinationFormatDescriptor.Id);
            Assert.True(report.Compatibility.IsStrictlyCompatible);
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void Convert_MacroDocumentToMacroFreeDocument_BlocksOrReportsRemoval() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docm");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Macro carrier");
                document.AddMacro(File.ReadAllBytes(Path.Combine(_directoryDocuments, "vbaProject.bin")));
                document.Save();
            }

            WordDocumentConversionException blocked = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(sourcePath, blockedPath));

            Assert.Equal(WordDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
            WordConversionDiagnostic finding = Assert.Single(blocked.Result.Report.Diagnostics,
                diagnostic => diagnostic.Code == "Word.VbaProject.Removed");
            Assert.Equal(OfficeCompatibilityState.Blocked, finding.CompatibilityState);
            Assert.True(finding.CompatibilityImpact.HasFlag(OfficeCompatibilityImpact.Security));
            Assert.False(File.Exists(blockedPath));

            WordDocumentConversionResult allowed = WordDocument.Convert(sourcePath, allowedPath,
                new WordDocumentConversionOptions { CompatibilityMode = OfficeCompatibilityMode.BestEffort });

            Assert.True(allowed.HasLoss);
            Assert.Equal(OfficeCompatibilityMode.BestEffort, allowed.Report.Compatibility.Mode);
            using WordDocument converted = WordDocument.Load(allowedPath);
            Assert.False(converted.HasMacros);
        }

        [Fact]
        public void LegacyDoc_Convert_DisablesOpenSettingsAutoSaveForSourceIsolation() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Source must remain untouched");
                document.Save(sourcePath);
            }
            File.SetLastWriteTimeUtc(sourcePath, new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc));
            byte[] sourceBytes = File.ReadAllBytes(sourcePath);
            DateTime sourceWriteTime = File.GetLastWriteTimeUtc(sourcePath);

            WordDocumentConversionException exception = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(sourcePath, destinationPath, new WordDocumentConversionOptions {
                    OpenSettings = new OpenSettings { AutoSave = true }
                }));

            Assert.Equal(WordDocumentConversionFailureReason.SameFormat, exception.Reason);
            Assert.Equal(sourceBytes, File.ReadAllBytes(sourcePath));
            Assert.Equal(sourceWriteTime, File.GetLastWriteTimeUtc(sourcePath));
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
            Assert.True(exception.Result.HasLoss);
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

        [Fact]
        public void LegacyDoc_Convert_ChartAndSmartArtUsePolicyDrivenVisualFallback() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string visualPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Executive scorecard");
                WordChart chart = document.AddChart("Delivery status", false, 420, 240);
                chart.AddPie("Complete", 72);
                chart.AddPie("Remaining", 28);
                WordSmartArt smartArt = document.AddSmartArt(SmartArtType.BasicProcess);
                while (smartArt.NodeCount < 3) smartArt.AddNode("Step " + smartArt.NodeCount);
                smartArt.ReplaceTexts("Plan", "Build", "Ship");
                document.Save();
            }

            WordDocumentConversionException blocked = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(sourcePath, blockedPath));

            Assert.Equal(WordDocumentConversionFailureReason.DestinationFeatureUnsupported, blocked.Reason);
            Assert.Contains(blocked.Result.Report.Diagnostics,
                finding => finding.Code == "Word.LegacyWriter.Unsupported"
                    && finding.CompatibilityState == OfficeCompatibilityState.Blocked);
            Assert.False(File.Exists(blockedPath));

            var options = new WordDocumentConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.PreferVisual
            };
            WordDocumentConversionReport preflight = WordDocument.AnalyzeConversion(sourcePath, visualPath, options);

            Assert.False(File.Exists(visualPath));
            Assert.Contains(preflight.Diagnostics,
                finding => finding.Code == "Word.LegacyWriter.VisualFallback"
                    && finding.CompatibilityState == OfficeCompatibilityState.Rasterized
                    && finding.FallbackArtifact != null);
            Assert.Contains(preflight.Diagnostics,
                finding => finding.Code == "Word.SourceCarrier.NotEmbedded"
                    && finding.CompatibilityState == OfficeCompatibilityState.Dropped);

            WordDocumentConversionResult converted = WordDocument.Convert(sourcePath, visualPath, options);

            Assert.Equal(
                preflight.Diagnostics.Select(finding => finding.Code),
                converted.Report.Diagnostics.Select(finding => finding.Code));
            Assert.True(converted.HasLoss);
            using WordDocument legacy = WordDocument.Load(visualPath);
            Assert.Equal(WordFileFormat.Doc, legacy.SourceFormat);
            Assert.NotEmpty(legacy.Images);
            Assert.Empty(legacy.Charts);
            Assert.Empty(legacy.SmartArts);
            Assert.False(legacy.TryGetCompatibilitySourcePayload(out _, out _));
        }

        [Fact]
        public void LegacyDoc_Convert_PreservationOnlyRetainsHashVerifiedOriginalSource() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Recoverable source");
                WordChart chart = document.AddChart("Preserved chart", false, 360, 200);
                chart.AddPie("Yes", 3);
                chart.AddPie("No", 1);
                document.Save();
            }
            byte[] sourceBytes = File.ReadAllBytes(sourcePath);

            WordDocumentConversionResult converted = WordDocument.Convert(
                sourcePath,
                destinationPath,
                new WordDocumentConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.PreservationOnly
                });

            WordConversionDiagnostic carrier = Assert.Single(converted.Report.Diagnostics,
                finding => finding.Code == "Word.SourceCarrier.Embedded");
            Assert.Equal(OfficeCompatibilityState.EmbeddedSource, carrier.CompatibilityState);
            Assert.False(carrier.RepresentsDataLoss);
            using WordDocument legacy = WordDocument.Load(destinationPath);
            Assert.True(legacy.TryGetCompatibilitySourcePayload(out OfficeCompatibilitySourcePayload? payload, out string? error), error);
            Assert.NotNull(payload);
            Assert.Equal("Word.Docx", payload!.FormatId);
            Assert.Equal(Path.GetFileName(sourcePath), payload.FileName);
            Assert.Equal(OfficeCompatibilityMode.PreservationOnly, payload.Mode);
            Assert.Equal(sourceBytes, payload.ToArray());
            Assert.Equal(sourceBytes.Length, payload.Length);
        }

        [Fact]
        public void LegacyDoc_Convert_PreservationOnlyRetainsSourceOnNativeLegacyPath() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Native and recoverable");
                document.Save();
            }
            byte[] sourceBytes = File.ReadAllBytes(sourcePath);

            WordDocumentConversionResult converted = WordDocument.Convert(
                sourcePath,
                destinationPath,
                new WordDocumentConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.PreservationOnly
                });

            Assert.Contains(converted.Report.Diagnostics,
                finding => finding.Code == "Word.SourceCarrier.Embedded"
                    && finding.CompatibilityState == OfficeCompatibilityState.EmbeddedSource);
            using WordDocument legacy = WordDocument.Load(destinationPath);
            Assert.True(legacy.TryGetCompatibilitySourcePayload(out OfficeCompatibilitySourcePayload? payload, out string? error), error);
            Assert.NotNull(payload);
            Assert.Equal(sourceBytes, payload!.ToArray());
        }

        [Fact]
        public void LegacyDoc_Convert_LegacyToModernPreservationOnlyRetainsImmediateSource() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string legacyPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Legacy source carrier");
                document.Save();
            }
            WordDocument.Convert(sourcePath, legacyPath).RequireNoLoss();
            byte[] legacyBytes = File.ReadAllBytes(legacyPath);

            WordDocumentConversionResult converted = WordDocument.Convert(
                legacyPath,
                destinationPath,
                new WordDocumentConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.PreservationOnly
                });

            Assert.Contains(converted.Report.Diagnostics,
                finding => finding.Code == "Word.SourceCarrier.Embedded"
                    && finding.CompatibilityState == OfficeCompatibilityState.EmbeddedSource);
            using WordDocument modern = WordDocument.Load(destinationPath);
            Assert.True(modern.TryGetCompatibilitySourcePayload(out OfficeCompatibilitySourcePayload? payload, out string? error), error);
            Assert.NotNull(payload);
            Assert.Equal("Word.Doc", payload!.FormatId);
            Assert.Equal(legacyBytes, payload.ToArray());
        }

        [Fact]
        public void LegacyDoc_Convert_EditableAndVisualModesBlockUnmappedActiveContent() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            File.WriteAllBytes(
                sourcePath,
                LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Active legacy source"));

            foreach (OfficeCompatibilityMode mode in new[] {
                         OfficeCompatibilityMode.PreferEditable,
                         OfficeCompatibilityMode.PreferVisual
                     }) {
                string destinationPath = Path.Combine(
                    _directoryWithFiles,
                    Guid.NewGuid().ToString("N") + ".docx");
                var options = new WordDocumentConversionOptions { CompatibilityMode = mode };

                WordDocumentConversionReport analysis = WordDocument.AnalyzeConversion(
                    sourcePath,
                    destinationPath,
                    options);

                Assert.True(analysis.Compatibility.HasBlockedFeatures);
                Assert.True(analysis.Compatibility.HasSecurityImpact);
                Assert.Contains(analysis.Compatibility.Findings,
                    finding => finding.State == OfficeCompatibilityState.Blocked
                        && finding.RepresentsLoss
                        && (finding.Impact & OfficeCompatibilityImpact.Security) != 0);
                WordDocumentConversionException blocked = Assert.Throws<WordDocumentConversionException>(() =>
                    WordDocument.Convert(sourcePath, destinationPath, options));
                Assert.Equal(WordDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
                Assert.False(File.Exists(destinationPath));
            }
        }

        [Fact]
        public void WordConversion_SignedModernSourceIsIncludedInAnalysisAndStructuredFailure() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".dotx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".dotx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Signed conversion source");
                document.Save();
            }
            AddDigitalSignatureMetadata(sourcePath, CreateSignatureXml());

            WordDocumentConversionReport analysis = WordDocument.AnalyzeConversion(sourcePath, blockedPath);
            Assert.True(analysis.Compatibility.HasSecurityImpact);
            Assert.Contains(analysis.Compatibility.Findings,
                finding => finding.Code == "Word.DigitalSignature.Invalidated"
                    && finding.State == OfficeCompatibilityState.Blocked
                    && finding.RepresentsLoss);

            WordDocumentConversionException blocked = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(sourcePath, blockedPath));
            Assert.Equal(WordDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
            Assert.False(File.Exists(blockedPath));

            WordDocumentConversionResult allowed = WordDocument.Convert(
                sourcePath,
                allowedPath,
                new WordDocumentConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                    SaveOptions = new WordSaveOptions {
                        SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation
                    }
                });
            Assert.True(allowed.Report.Compatibility.HasSecurityImpact);
            Assert.Contains(allowed.Report.Compatibility.Findings,
                finding => finding.Code == "Word.DigitalSignature.Invalidated"
                    && finding.State == OfficeCompatibilityState.Dropped
                    && finding.RepresentsLoss);
            Assert.True(File.Exists(allowedPath));
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
