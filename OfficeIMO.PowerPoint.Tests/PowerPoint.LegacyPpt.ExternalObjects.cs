using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        public enum ExternalObjectFixtureKind {
            LinkedOle,
            ActiveX,
            EmbeddedWaveMedia,
            LinkedWaveMedia,
            MidiAudio,
            AviMovie,
            MciMovie,
            CdAudio
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void LinkedOleObject_ImportsTypedAndRemainsExactAcrossUnrelatedEdit(
            bool compressed) {
            byte[] storageBytes = CreateOleTestStorage(
                compressed ? "Compressed linked OLE" : "Linked OLE");
            byte[] sourceBytes = CreateExternalObjectFixture(storageBytes,
                ExternalObjectFixtureKind.LinkedOle, compressed);
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                sourceBytes);

            LegacyPptLinkedOleObject linked = Assert.Single(
                source.LinkedOleObjects);
            Assert.Empty(source.OleObjects);
            Assert.Empty(source.ActiveXControls);
            Assert.Equal(LegacyPptOleUpdateMode.Manual,
                linked.UpdateMode);
            Assert.Equal("Package", linked.ProgId);
            Assert.Equal(storageBytes, linked.GetBytes());
            Assert.Equal(compressed, linked.WasCompressed);
            Assert.False(linked.HasMetafile);
            Assert.Equal(0, linked.MetafileByteCount);
            Assert.Null(linked.GetMetafileRecordBytes());
            LegacyPptShape sourceShape = Assert.Single(source.Slides[0]
                .Shapes, shape => shape.LinkedOleObject != null);
            Assert.Equal(LegacyPptShapeKind.Unsupported,
                sourceShape.Kind);
            Assert.Same(linked, sourceShape.LinkedOleObject);
            Assert.Equal(source.Slides[0].SlideId, linked.SlideId);
            Assert.Contains(source.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-OLE-LINK-PRESERVED");

            LegacyPptImportReport report = source.CreateImportReport();
            Assert.Equal(1, report.LinkedOleObjectCount);
            Assert.Equal(storageBytes.Length,
                report.LinkedOleObjectByteCount);
            Assert.Equal(compressed ? 1 : 0,
                report.CompressedLinkedOleObjectCount);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                PowerPointTextBox textBox = Assert.Single(
                    imported.Slides[0].TextBoxes, item =>
                        item.Text == "Editable companion");
                Assert.Equal("Editable companion", textBox.Text);
                PowerPointFeatureFinding finding = Assert.Single(imported
                    .InspectFeatures().FindFeatures("Linked OLE objects"));
                Assert.Equal(PowerPointFeatureSupportLevel.Preserved,
                    finding.SupportLevel);
                Assert.Equal(1, finding.Count);
                LegacyPptWritePreflightReport exactPreflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.False(exactPreflight.CanWrite);
                Assert.Contains(exactPreflight.Findings, item =>
                    item.Code == "PPT-WRITE-PRESERVED-LINKED-OLE");
                Assert.Throws<NotSupportedException>(() =>
                    imported.ToBytes(PowerPointFileFormat.Ppt));
                Assert.Equal(sourceBytes, imported.ToBytes(
                    PowerPointFileFormat.Ppt, new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    }));
                textBox.Text = "Edited companion";
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    });
            }

            AssertExternalObjectPreserved(source, savedBytes,
                linked.PersistId, storageBytes,
                ExternalObjectFixtureKind.LinkedOle, compressed);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ActiveXControl_ImportsTypedAndRemainsExactAcrossUnrelatedEdit(
            bool compressed) {
            byte[] storageBytes = CreateOleTestStorage(
                compressed ? "Compressed ActiveX" : "ActiveX");
            byte[] sourceBytes = CreateExternalObjectFixture(storageBytes,
                ExternalObjectFixtureKind.ActiveX, compressed);
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                sourceBytes);

            LegacyPptActiveXControl control = Assert.Single(
                source.ActiveXControls);
            Assert.Empty(source.OleObjects);
            Assert.Empty(source.LinkedOleObjects);
            Assert.Equal("Package", control.ProgId);
            Assert.Equal(storageBytes, control.GetBytes());
            Assert.Equal(compressed, control.WasCompressed);
            Assert.False(control.HasMetafile);
            Assert.Equal(0, control.MetafileByteCount);
            Assert.Null(control.GetMetafileRecordBytes());
            LegacyPptShape sourceShape = Assert.Single(source.Slides[0]
                .Shapes, shape => shape.ActiveXControl != null);
            Assert.Equal(LegacyPptShapeKind.Unsupported,
                sourceShape.Kind);
            Assert.Same(control, sourceShape.ActiveXControl);
            Assert.Equal(source.Slides[0].SlideId, control.SlideId);
            Assert.Contains(source.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-ACTIVEX-PRESERVED");

            LegacyPptImportReport report = source.CreateImportReport();
            Assert.Equal(1, report.ActiveXControlCount);
            Assert.Equal(storageBytes.Length,
                report.ActiveXControlByteCount);
            Assert.Equal(compressed ? 1 : 0,
                report.CompressedActiveXControlCount);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                PowerPointTextBox textBox = Assert.Single(
                    imported.Slides[0].TextBoxes, item =>
                        item.Text == "Editable companion");
                Assert.Equal("Editable companion", textBox.Text);
                PowerPointFeatureFinding finding = Assert.Single(imported
                    .InspectFeatures().FindFeatures("ActiveX controls"));
                Assert.Equal(PowerPointFeatureSupportLevel.Preserved,
                    finding.SupportLevel);
                Assert.Equal(1, finding.Count);
                LegacyPptWritePreflightReport exactPreflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.False(exactPreflight.CanWrite);
                Assert.Contains(exactPreflight.Findings, item =>
                    item.Code == "PPT-WRITE-PRESERVED-ACTIVEX");
                Assert.Throws<NotSupportedException>(() =>
                    imported.ToBytes(PowerPointFileFormat.Ppt));
                Assert.Equal(sourceBytes, imported.ToBytes(
                    PowerPointFileFormat.Ppt, new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    }));
                textBox.Text = "Edited companion";
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    });
            }

            AssertExternalObjectPreserved(source, savedBytes,
                control.PersistId, storageBytes,
                ExternalObjectFixtureKind.ActiveX, compressed);
        }

        [Fact]
        public void CapabilityContract_ReportsLinkedOleAndActiveXPreservation() {
            LegacyPptCapability linked = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.LinkedOle);
            LegacyPptCapability activeX = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.ActiveX);

            Assert.Equal(LegacyPptCapabilityState.Preserved,
                linked.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                linked.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                linked.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                linked.PptxToBinary);
            Assert.Contains("exact compound-storage bytes", linked.Note);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                activeX.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                activeX.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                activeX.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                activeX.PptxToBinary);
            Assert.Contains("exact Office Forms storage bytes",
                activeX.Note);
        }

        [Fact]
        public void MalformedLinkedOleMetadata_IsDiagnosedAndRemainsOpaque() {
            byte[] storageBytes = CreateOleTestStorage(
                "Malformed linked OLE metadata");
            byte[] sourceBytes = CreateExternalObjectFixture(storageBytes,
                ExternalObjectFixtureKind.LinkedOle, compressed: false,
                linkedUpdateMode: 99);

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                sourceBytes);
            Assert.Empty(neutral.LinkedOleObjects);
            Assert.Contains(neutral.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-OLE-LINK-IDENTITY");
            Assert.Contains(neutral.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-OLE-SHAPE-TARGET");
            Assert.True(neutral.CreateImportReport().HasConversionLoss);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            Assert.Contains(projected.LegacyPptImportDiagnostics,
                diagnostic => diagnostic.Code ==
                    "PPT-OLE-LINK-IDENTITY");
            LegacyPptWritePreflightReport preflight = projected
                .AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-PRESERVED-LINKED-OLE");
            Assert.Throws<NotSupportedException>(() =>
                projected.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Equal(sourceBytes, projected.ToBytes(
                PowerPointFileFormat.Ppt, new PowerPointSaveOptions {
                    LossPolicy = PowerPointConversionLossPolicy.Allow
                }));
        }

        private static void AssertExternalObjectPreserved(
            LegacyPptPresentation source, byte[] savedBytes,
            uint persistId, byte[] storageBytes,
            ExternalObjectFixtureKind kind, bool compressed) {
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            byte[] savedStorage;
            bool savedCompressed;
            if (kind == ExternalObjectFixtureKind.LinkedOle) {
                LegacyPptLinkedOleObject linked = Assert.Single(
                    saved.LinkedOleObjects);
                savedStorage = linked.GetBytes();
                savedCompressed = linked.WasCompressed;
            } else {
                LegacyPptActiveXControl control = Assert.Single(
                    saved.ActiveXControls);
                savedStorage = control.GetBytes();
                savedCompressed = control.WasCompressed;
            }
            Assert.Equal(storageBytes, savedStorage);
            Assert.Equal(compressed, savedCompressed);
            Assert.Equal(source.Package.PersistObjects[persistId]
                    .RecordBytes,
                saved.Package.PersistObjects[persistId].RecordBytes);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));
            Assert.Contains(saved.Slides[0].Shapes,
                shape => shape.Text == "Edited companion");
        }

        private static byte[] CreateExternalObjectFixture(
            byte[] storageBytes, ExternalObjectFixtureKind kind,
            bool compressed, uint linkedUpdateMode = 3) {
            byte[] embeddedBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide();
                slide.AddTextBox("Editable companion");
                using var storage = new MemoryStream(storageBytes,
                    writable: false);
                slide.AddOleObject(storage, "Package");
                embeddedBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            if (compressed) {
                embeddedBytes = ConvertOleStorageToCompressed(
                    embeddedBytes);
            }
            return ConvertEmbeddedOleContainer(embeddedBytes, kind,
                linkedUpdateMode);
        }

        private static byte[] ConvertEmbeddedOleContainer(
            byte[] sourceBytes, ExternalObjectFixtureKind kind,
            uint linkedUpdateMode, uint embeddedSoundId = 0,
            ushort mediaFlags = 0) {
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptPersistObject documentPersist = source.Package
                .PersistObjects[source.Package.DocumentPersistId];
            LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(
                documentPersist.RecordBytes, 0,
                new LegacyPptImportOptions());
            int convertedCount = 0;
            byte[] rewrittenDocument = RewriteExternalObjectRecord(document,
                kind, source.Slides[0].SlideId, linkedUpdateMode,
                embeddedSoundId, mediaFlags, ref convertedCount);
            Assert.Equal(1, convertedCount);

            var offsets = source.Package.PersistObjectOffsets.ToDictionary(
                pair => pair.Key, pair => pair.Value);
            using var stream = new MemoryStream();
            stream.Write(source.Package.DocumentStream, 0,
                source.Package.DocumentStream.Length);
            offsets[source.Package.DocumentPersistId] = checked(
                (uint)stream.Position);
            stream.Write(rewrittenDocument, 0, rewrittenDocument.Length);
            uint directoryOffset = checked((uint)stream.Position);
            byte[] directory = BuildVbaPersistDirectory(offsets);
            stream.Write(directory, 0, directory.Length);
            uint editOffset = checked((uint)stream.Position);
            int oldEditOffset = checked((int)source.Package
                .CurrentEditOffset);
            int editLength = checked(8 + (int)ReadVbaUInt32(
                source.Package.DocumentStream, oldEditOffset + 4));
            var edit = new byte[editLength];
            Buffer.BlockCopy(source.Package.DocumentStream, oldEditOffset,
                edit, 0, edit.Length);
            WriteVbaUInt32(edit, 20, directoryOffset);
            stream.Write(edit, 0, edit.Length);

            byte[] currentUser = (byte[])source.Package.CurrentUserStream
                .Clone();
            WriteVbaUInt32(currentUser, 16, editOffset);
            return source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(
                    StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = stream.ToArray(),
                    ["Current User"] = currentUser
                });
        }

        private static byte[] RewriteExternalObjectRecord(
            LegacyPptRecord record, ExternalObjectFixtureKind kind,
            uint slideId, uint linkedUpdateMode,
            uint embeddedSoundId, ushort mediaFlags,
            ref int convertedCount) {
            if (record.Type == 0x0FCC) {
                convertedCount++;
                if (kind >= ExternalObjectFixtureKind.EmbeddedWaveMedia) {
                    LegacyPptRecord objectAtom = Assert.Single(
                        record.Children, child => child.Type == 0x0FC3);
                    uint id = objectAtom.ReadUInt32(8);
                    var mediaPayload = new byte[8];
                    WriteVbaUInt32(mediaPayload, 0, id);
                    WriteVbaUInt16(mediaPayload, 4, mediaFlags);
                    var mediaChildren = new List<byte[]> {
                        BuildVbaRecord(version: 0, instance: 0,
                            type: 0x1004, mediaPayload)
                    };
                    ushort containerType;
                    switch (kind) {
                        case ExternalObjectFixtureKind.EmbeddedWaveMedia:
                            var wavePayload = new byte[8];
                            WriteVbaUInt32(wavePayload, 0,
                                embeddedSoundId);
                            WriteVbaUInt32(wavePayload, 4, 2500);
                            mediaChildren.Add(BuildVbaRecord(version: 1,
                                instance: 1, type: 0x1013,
                                wavePayload));
                            containerType = 0x100F;
                            break;
                        case ExternalObjectFixtureKind.LinkedWaveMedia:
                            mediaChildren.Add(BuildMediaPathRecord(
                                @"C:\Media\sample.wav"));
                            containerType = 0x1010;
                            break;
                        case ExternalObjectFixtureKind.MidiAudio:
                            mediaChildren.Add(BuildMediaPathRecord(
                                @"C:\Media\sample.mid"));
                            containerType = 0x100D;
                            break;
                        case ExternalObjectFixtureKind.CdAudio:
                            var cdPayload = new byte[8];
                            WriteVbaUInt32(cdPayload, 0, 0x01020304);
                            WriteVbaUInt32(cdPayload, 4, 0x05060708);
                            mediaChildren.Add(BuildVbaRecord(version: 0,
                                instance: 0, type: 0x1012,
                                cdPayload));
                            containerType = 0x100E;
                            break;
                        case ExternalObjectFixtureKind.AviMovie:
                        case ExternalObjectFixtureKind.MciMovie:
                            mediaChildren.Add(BuildMediaPathRecord(
                                @"C:\Media\sample.avi"));
                            byte[] video = BuildVbaRecord(version: 0x0F,
                                instance: 0, type: 0x1005,
                                JoinExternalObjectRecords(mediaChildren));
                            containerType = kind ==
                                ExternalObjectFixtureKind.AviMovie
                                ? (ushort)0x1006
                                : (ushort)0x1007;
                            return BuildVbaRecord(version: 0x0F,
                                instance: 0, type: containerType,
                                video);
                        default:
                            throw new InvalidOperationException(
                                "Unsupported media fixture kind.");
                    }
                    return BuildVbaRecord(version: 0x0F, instance: 0,
                        type: containerType,
                        JoinExternalObjectRecords(mediaChildren));
                }
                var children = new List<byte[]>(record.Children.Count);
                foreach (LegacyPptRecord child in record.Children) {
                    if (child.Type == 0x0FCD) {
                        var payload = new byte[kind ==
                            ExternalObjectFixtureKind.LinkedOle ? 12 : 4];
                        WriteVbaUInt32(payload, 0, slideId);
                        if (kind == ExternalObjectFixtureKind.LinkedOle) {
                            WriteVbaUInt32(payload, 4, linkedUpdateMode);
                        }
                        children.Add(BuildVbaRecord(version: 0,
                            instance: 0, type: kind ==
                                ExternalObjectFixtureKind.LinkedOle
                                ? (ushort)0x0FD1
                                : (ushort)0x0FFB,
                            payload));
                    } else if (child.Type == 0x0FC3) {
                        byte[] objectAtom = child.CopyRecordBytes();
                        WriteVbaUInt32(objectAtom, 12,
                            kind == ExternalObjectFixtureKind.LinkedOle
                                ? 1U
                                : 2U);
                        children.Add(objectAtom);
                    } else {
                        children.Add(child.CopyRecordBytes());
                    }
                }
                return BuildVbaRecord(version: 0x0F, instance: 0,
                    type: kind == ExternalObjectFixtureKind.LinkedOle
                        ? (ushort)0x0FCE
                        : (ushort)0x0FEE,
                    JoinExternalObjectRecords(children));
            }
            if (record.Version != 0x0F) return record.CopyRecordBytes();
            var rewrittenChildren = new List<byte[]>(record.Children.Count);
            foreach (LegacyPptRecord child in record.Children) {
                rewrittenChildren.Add(RewriteExternalObjectRecord(child,
                    kind, slideId, linkedUpdateMode,
                    embeddedSoundId, mediaFlags,
                    ref convertedCount));
            }
            return BuildVbaRecord(record.Version, record.Instance,
                record.Type, JoinExternalObjectRecords(rewrittenChildren));
        }

        private static byte[] JoinExternalObjectRecords(
            IEnumerable<byte[]> records) {
            using var output = new MemoryStream();
            foreach (byte[] record in records) {
                output.Write(record, 0, record.Length);
            }
            return output.ToArray();
        }

        private static byte[] BuildMediaPathRecord(string path) =>
            BuildVbaRecord(version: 0, instance: 0, type: 0x0FBA,
                System.Text.Encoding.Unicode.GetBytes(path));
    }
}
