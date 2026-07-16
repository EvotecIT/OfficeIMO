using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OpenMcdf;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;
using CfbVersion = OpenMcdf.Version;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void FreshEmbeddedOleObject_WritesNativeStorageAndProjectsToEditableModel() {
            byte[] storageBytes = CreateOleTestStorage("Fresh OLE payload");
            byte[] binary;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide(
                    P.SlideLayoutValues.Blank);
                using var storage = new MemoryStream(storageBytes,
                    writable: false);
                PowerPointOleObject ole = slide.AddOleObject(storage,
                    "Package", 12700L, 25400L, 2743200L, 1828800L);
                ole.ShowAsIcon = true;
                ole.FollowColorScheme =
                    P.OleObjectFollowColorSchemeValues.TextAndBackground;

                Assert.Equal(PowerPointShapeContentType.OleObject,
                    ole.ShapeContentType);
                Assert.Equal("Package", ole.ProgId);
                Assert.Equal(storageBytes, ole.GetData());
                Assert.Empty(created.ValidateDocument());

                LegacyPptWritePreflightReport preflight = created
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                binary = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(binary);
            LegacyPptEmbeddedOleObject embedded = Assert.Single(
                neutral.OleObjects);
            Assert.False(embedded.WasCompressed);
            Assert.Equal(LegacyPptOleDrawAspect.Icon,
                embedded.DrawAspect);
            Assert.Equal(LegacyPptOleColorFollow.TextAndBackground,
                embedded.ColorFollow);
            Assert.Equal("Package", embedded.ProgId);
            Assert.Equal(storageBytes, embedded.GetBytes());
            Assert.Equal(1, neutral.CreateImportReport()
                .EmbeddedOleObjectCount);
            Assert.Equal(storageBytes.Length, neutral.CreateImportReport()
                .EmbeddedOleObjectByteCount);
            Assert.Equal(0, neutral.CreateImportReport()
                .CompressedEmbeddedOleObjectCount);
            Assert.Equal(0x1011, neutral.Package.PersistObjects[
                embedded.PersistId].RecordType);
            LegacyPptShape shape = Assert.Single(neutral.Slides[0].Shapes);
            Assert.Equal(LegacyPptShapeKind.OleObject, shape.Kind);
            Assert.Same(embedded, shape.OleObject);
            Assert.DoesNotContain(neutral.Diagnostics, diagnostic =>
                diagnostic.Code.StartsWith("PPT-OLE-",
                    StringComparison.Ordinal));

            using var input = new MemoryStream(binary);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            PowerPointOleObject projectedOle = Assert.IsType<
                PowerPointOleObject>(Assert.Single(projected.Slides[0].Shapes));
            Assert.Equal("Package", projectedOle.ProgId);
            Assert.True(projectedOle.ShowAsIcon);
            Assert.Equal(P.OleObjectFollowColorSchemeValues.TextAndBackground,
                projectedOle.FollowColorScheme);
            Assert.Equal(storageBytes, projectedOle.GetData());
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(binary,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void EmbeddedOleObjectAndVbaProject_UseDistinctNativePersistObjects() {
            byte[] oleStorage = CreateOleTestStorage("OLE and VBA");
            byte[] vbaProject = CreateVbaTestProject("OleVbaModule",
                "Sub OleVba()\nEnd Sub");
            byte[] binary;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide(
                    P.SlideLayoutValues.Blank);
                using var storage = new MemoryStream(oleStorage,
                    writable: false);
                slide.AddOleObject(storage, "Package");
                SetVbaProject(created, vbaProject);
                Assert.Empty(created.ValidateDocument());
                binary = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(binary);
            LegacyPptEmbeddedOleObject ole = Assert.Single(neutral.OleObjects);
            LegacyPptVbaProject vba = Assert.IsType<LegacyPptVbaProject>(
                neutral.VbaProject);
            Assert.NotEqual(ole.PersistId, vba.PersistId);
            Assert.Equal(oleStorage, ole.GetBytes());
            Assert.Equal(vbaProject, vba.GetBytes());

            using var input = new MemoryStream(binary);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            Assert.Equal(oleStorage, Assert.IsType<PowerPointOleObject>(
                Assert.Single(projected.Slides[0].Shapes)).GetData());
            Assert.Equal(vbaProject, ReadVbaProject(projected));
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void ImportedEmbeddedOleObject_StorageMetadataAndGeometryUseIncrementalRecords() {
            byte[] originalStorage = CreateOleTestStorage("Original OLE");
            byte[] replacementStorage = CreateOleTestStorage(
                "Replacement OLE");
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide(
                    P.SlideLayoutValues.Blank);
                using var storage = new MemoryStream(originalStorage,
                    writable: false);
                PowerPointOleObject ole = slide.AddOleObject(storage,
                    "Package", 12700L, 25400L, 2743200L, 1828800L);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                PowerPointOleObject ole = Assert.IsType<PowerPointOleObject>(
                    Assert.Single(imported.Slides[0].Shapes));
                using var replacement = new MemoryStream(replacementStorage,
                    writable: false);
                ole.UpdateData(replacement);
                ole.ProgId = "Word.Document.8";
                ole.ShowAsIcon = true;
                ole.FollowColorScheme =
                    P.OleObjectFollowColorSchemeValues.Full;
                ole.Left += 15875L;

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptEmbeddedOleObject embedded = Assert.Single(
                saved.OleObjects);
            Assert.Equal(replacementStorage, embedded.GetBytes());
            Assert.Equal("Word.Document.8", embedded.ProgId);
            Assert.Equal(2U, embedded.SubType);
            Assert.Equal(LegacyPptOleDrawAspect.Icon,
                embedded.DrawAspect);
            Assert.Equal(LegacyPptOleColorFollow.Scheme,
                embedded.ColorFollow);
            Assert.Equal(source.Slides[0].Shapes[0].Bounds.Left + 10,
                saved.Slides[0].Shapes[0].Bounds.Left);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));

            using var savedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(savedInput);
            PowerPointOleObject projectedOle = Assert.IsType<
                PowerPointOleObject>(Assert.Single(projected.Slides[0].Shapes));
            Assert.Equal(replacementStorage, projectedOle.GetData());
            Assert.Equal("Word.Document.8", projectedOle.ProgId);
            Assert.True(projectedOle.ShowAsIcon);
            Assert.Equal(P.OleObjectFollowColorSchemeValues.Full,
                projectedOle.FollowColorScheme);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void EmbeddedOleObject_RejectsInvalidReplacementWithoutChangingPart() {
            byte[] storageBytes = CreateOleTestStorage("Valid OLE");
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            using var storage = new MemoryStream(storageBytes,
                writable: false);
            PowerPointOleObject ole = slide.AddOleObject(storage, "Package");

            Assert.Throws<InvalidDataException>(() => ole.UpdateData(
                new MemoryStream(new byte[] { 1, 2, 3, 4 })));
            Assert.Equal(storageBytes, ole.GetData());
        }

        [Fact]
        public void EmbeddedOleObject_DuplicateOwnsIndependentStorageAndRemovalCleansPart() {
            byte[] originalStorage = CreateOleTestStorage("Original part");
            byte[] duplicateStorage = CreateOleTestStorage("Duplicate part");
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            using var source = new MemoryStream(originalStorage,
                writable: false);
            PowerPointOleObject original = slide.AddOleObject(source,
                "Package");

            PowerPointOleObject duplicate = Assert.IsType<PowerPointOleObject>(
                original.Duplicate(12700L, 12700L));
            Assert.Equal(2, slide.OleObjects.Count());
            Assert.Equal(originalStorage, duplicate.GetData());
            using (var replacement = new MemoryStream(duplicateStorage,
                       writable: false)) {
                duplicate.UpdateData(replacement);
            }
            Assert.Equal(originalStorage, original.GetData());
            Assert.Equal(duplicateStorage, duplicate.GetData());
            Assert.Equal(2, presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Single().GetPartsOfType<EmbeddedObjectPart>()
                .Count());

            original.Remove();

            Assert.Same(duplicate, Assert.Single(slide.OleObjects));
            Assert.Equal(1, presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Single().GetPartsOfType<EmbeddedObjectPart>()
                .Count());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void CompressedEmbeddedOleObject_ImportsAndRemainsExactAcrossUnrelatedEdit() {
            byte[] storageBytes = CreateOleTestStorage("Compressed OLE");
            byte[] uncompressedBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide(
                    P.SlideLayoutValues.Blank);
                using var storage = new MemoryStream(storageBytes,
                    writable: false);
                slide.AddOleObject(storage, "Package");
                uncompressedBytes = created.ToBytes(
                    PowerPointFileFormat.Ppt);
            }
            byte[] sourceBytes = ConvertOleStorageToCompressed(
                uncompressedBytes);
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptEmbeddedOleObject sourceOle = Assert.Single(
                source.OleObjects);
            Assert.True(sourceOle.WasCompressed);
            Assert.Equal(storageBytes, sourceOle.GetBytes());
            Assert.Equal(1, source.CreateImportReport()
                .CompressedEmbeddedOleObjectCount);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                Assert.Equal(sourceBytes,
                    imported.ToBytes(PowerPointFileFormat.Ppt));
                PowerPointOleObject ole = Assert.Single(
                    imported.Slides[0].OleObjects);
                ole.Left += 15875L;
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptEmbeddedOleObject savedOle = Assert.Single(
                saved.OleObjects);
            Assert.True(savedOle.WasCompressed);
            Assert.Equal(storageBytes, savedOle.GetBytes());
            Assert.Equal(source.Package.PersistObjects[sourceOle.PersistId]
                    .RecordBytes,
                saved.Package.PersistObjects[savedOle.PersistId]
                    .RecordBytes);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));
        }

        [Fact]
        public void FeatureReport_ReportsReferencedEmbeddedOleAsEditable() {
            byte[] storageBytes = CreateOleTestStorage("Feature report");
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            using var storage = new MemoryStream(storageBytes,
                writable: false);
            slide.AddOleObject(storage, "Package");

            PowerPointFeatureReport report = presentation.InspectFeatures();
            PowerPointFeatureFinding finding = Assert.Single(
                report.FindFeatures("Embedded OLE objects"));
            Assert.Equal(PowerPointFeatureSupportLevel.Editable,
                finding.SupportLevel);
            Assert.Equal(1, finding.Count);
            Assert.Empty(report.FindFeatures("Embedded packages"));
            report.EnsureNoAdvancedFeatures();
        }

        [Fact]
        public void CapabilityContract_ReportsEmbeddedOleParity() {
            LegacyPptCapability capability = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.EmbeddedOle);

            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                capability.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.PptxToBinary);
            Assert.Contains("compressed or uncompressed", capability.Note);
            Assert.Contains("loss-blocked", capability.Note);
        }

        private static byte[] CreateOleTestStorage(string contents) {
            using var output = new MemoryStream();
            using (RootStorage root = RootStorage.Create(output,
                       CfbVersion.V3, StorageModeFlags.LeaveOpen)) {
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(contents);
                using (CfbStream native = root.CreateStream(
                           "\u0001Ole10Native")) {
                    native.Write(bytes, 0, bytes.Length);
                }
                using (CfbStream stream = root.CreateStream("CONTENTS")) {
                    stream.Write(bytes, 0, bytes.Length);
                }
            }
            return output.ToArray();
        }

        private static byte[] ConvertOleStorageToCompressed(
            byte[] sourceBytes) {
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptEmbeddedOleObject ole = Assert.Single(
                source.OleObjects);
            var persist = source.Package.PersistObjects[ole.PersistId];
            Assert.Equal(source.Package.PersistObjectOffsets.Values.Max(),
                persist.StreamOffset);

            byte[] compressed = CompressVbaZlib(ole.GetBytes());
            var payload = new byte[4 + compressed.Length];
            WriteVbaUInt32(payload, 0, checked((uint)ole.Length));
            Buffer.BlockCopy(compressed, 0, payload, 4,
                compressed.Length);
            byte[] storage = BuildVbaRecord(version: 0, instance: 1,
                type: 0x1011, payload);

            using var document = new MemoryStream();
            document.Write(source.Package.DocumentStream, 0,
                checked((int)persist.StreamOffset));
            document.Write(storage, 0, storage.Length);
            uint directoryOffset = checked((uint)document.Position);
            byte[] directory = BuildVbaPersistDirectory(
                source.Package.PersistObjectOffsets);
            document.Write(directory, 0, directory.Length);
            uint editOffset = checked((uint)document.Position);
            int oldEditOffset = checked((int)source.Package
                .CurrentEditOffset);
            int editLength = checked(8 + (int)ReadVbaUInt32(
                source.Package.DocumentStream, oldEditOffset + 4));
            var edit = new byte[editLength];
            Buffer.BlockCopy(source.Package.DocumentStream, oldEditOffset,
                edit, 0, edit.Length);
            WriteVbaUInt32(edit, 20, directoryOffset);
            document.Write(edit, 0, edit.Length);

            byte[] currentUser = (byte[])source.Package.CurrentUserStream
                .Clone();
            WriteVbaUInt32(currentUser, 16, editOffset);
            return source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(
                    StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = document.ToArray(),
                    ["Current User"] = currentUser
                });
        }
    }
}
