using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OpenMcdf;
using Xunit;
using CfbVersion = OpenMcdf.Version;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void FreshBinaryVbaProject_WritesNativeStorageAndProjectsToMacroEnabledModel() {
            byte[] projectBytes = CreateVbaTestProject("FreshModule",
                "Sub Fresh()\nEnd Sub");
            byte[] binary;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide().AddTitle("VBA");
                SetVbaProject(created, projectBytes);

                LegacyPptWritePreflightReport preflight = created.AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                binary = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(binary);
            LegacyPptVbaProject vba = Assert.IsType<LegacyPptVbaProject>(
                neutral.VbaProject);
            Assert.False(vba.WasCompressed);
            Assert.Equal(projectBytes, vba.GetBytes());
            Assert.Equal(1, neutral.CreateImportReport().VbaProjectCount);
            Assert.Equal(projectBytes.Length,
                neutral.CreateImportReport().VbaProjectByteCount);
            Assert.False(neutral.CreateImportReport()
                .VbaProjectWasCompressed);
            Assert.Equal(0x1011, neutral.Package.PersistObjects[
                vba.PersistId].RecordType);
            Assert.DoesNotContain(neutral.Diagnostics, diagnostic =>
                diagnostic.Code.StartsWith("PPT-VBA-", StringComparison.Ordinal));

            using var input = new MemoryStream(binary);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            Assert.Equal(PresentationDocumentType.MacroEnabledPresentation,
                projected.OpenXmlDocument.DocumentType);
            Assert.Equal(projectBytes, ReadVbaProject(projected));
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(binary, projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void PptxOutput_StripsVbaWithoutMutatingBinaryModel() {
            byte[] projectBytes = CreateVbaTestProject("PptxModule",
                "Sub PptxExport()\nEnd Sub");
            byte[] binary;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide().AddTitle("Macro source");
                SetVbaProject(created, projectBytes);
                binary = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(binary);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            byte[] pptx = imported.ToBytes(PowerPointFileFormat.Pptx);

            using (var package = new MemoryStream(pptx, writable: false))
            using (PresentationDocument document =
                   PresentationDocument.Open(package, false)) {
                Assert.Equal(PresentationDocumentType.Presentation,
                    document.DocumentType);
                Assert.Null(document.PresentationPart!.VbaProjectPart);
            }
            Assert.Equal(projectBytes, ReadVbaProject(imported));
            byte[] binaryAfterPptxExport = imported.ToBytes(
                PowerPointFileFormat.Ppt);
            Assert.Equal(projectBytes, Assert.IsType<LegacyPptVbaProject>(
                LegacyPptPresentation.Load(binaryAfterPptxExport)
                    .VbaProject).GetBytes());
        }

        [Fact]
        public void ImportedVbaProject_ReplacementAndRemovalUseIncrementalPersistRecords() {
            byte[] originalProject = CreateVbaTestProject("OriginalModule",
                "Sub Original()\nEnd Sub");
            byte[] replacementProject = CreateVbaTestProject(
                "ReplacementModule", "Sub Replacement()\nEnd Sub");
            byte[] sourceBytes;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide().AddTitle("Replace VBA");
                SetVbaProject(created, originalProject);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation source = LegacyPptPresentation.Load(sourceBytes);

            byte[] replacedBytes;
            using (var sourceStream = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(
                       sourceStream)) {
                SetVbaProject(imported, replacementProject);
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                replacedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation replaced = LegacyPptPresentation.Load(
                replacedBytes);
            Assert.Equal(replacementProject,
                Assert.IsType<LegacyPptVbaProject>(replaced.VbaProject)
                    .GetBytes());
            Assert.True(replaced.Package.DocumentStream.AsSpan(0,
                    source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));

            byte[] removedBytes;
            using (var replacedStream = new MemoryStream(replacedBytes))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(
                       replacedStream)) {
                PresentationPart presentationPart = imported.OpenXmlDocument
                    .PresentationPart!;
                presentationPart.DeletePart(presentationPart.VbaProjectPart!);
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation removed = LegacyPptPresentation.Load(
                removedBytes);
            Assert.Null(removed.VbaProject);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    replaced.Package.DocumentStream.Length)
                .SequenceEqual(replaced.Package.DocumentStream));
        }

        [Fact]
        public void ImportedPresentation_CanAddVbaProjectWithoutRebuildingSourceGraph() {
            byte[] sourceBytes = File.ReadAllBytes(FixturePath);
            LegacyPptPresentation source = LegacyPptPresentation.Load(sourceBytes);
            byte[] projectBytes = CreateVbaTestProject("AddedModule",
                "Sub Added()\nEnd Sub");

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                SetVbaProject(imported, projectBytes);
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Equal(projectBytes,
                Assert.IsType<LegacyPptVbaProject>(saved.VbaProject)
                    .GetBytes());
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));
            Assert.Equal(source.Package.CopyCompoundStreams()
                    .Where(pair => pair.Key != "PowerPoint Document"
                                   && pair.Key != "Current User")
                    .Select(pair => pair.Key)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase),
                saved.Package.CopyCompoundStreams()
                    .Where(pair => pair.Key != "PowerPoint Document"
                                   && pair.Key != "Current User")
                    .Select(pair => pair.Key)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void CompressedBinaryVbaProject_ImportsAndRemainsExactAcrossUnrelatedEdit() {
            byte[] projectBytes = CreateVbaTestProject("CompressedModule",
                "Sub Compressed()\nEnd Sub");
            byte[] uncompressedBytes;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide().AddTitle("Compressed VBA");
                SetVbaProject(created, projectBytes);
                uncompressedBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            byte[] sourceBytes = ConvertVbaStorageToCompressed(
                uncompressedBytes);
            LegacyPptPresentation source = LegacyPptPresentation.Load(sourceBytes);
            LegacyPptVbaProject sourceVba = Assert.IsType<LegacyPptVbaProject>(
                source.VbaProject);
            Assert.True(sourceVba.WasCompressed);
            Assert.Equal(projectBytes, sourceVba.GetBytes());
            Assert.True(source.CreateImportReport()
                .VbaProjectWasCompressed);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                Assert.Equal(projectBytes, ReadVbaProject(imported));
                Assert.Equal(sourceBytes,
                    imported.ToBytes(PowerPointFileFormat.Ppt));
                PowerPointTextBox title = Assert.Single(imported.Slides[0]
                    .TextBoxes, item => item.Text == "Compressed VBA");
                title.Left += 15875L;
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptVbaProject savedVba = Assert.IsType<LegacyPptVbaProject>(
                saved.VbaProject);
            Assert.True(savedVba.WasCompressed);
            Assert.Equal(projectBytes, savedVba.GetBytes());
            Assert.Equal(source.Package.PersistObjects[sourceVba.PersistId]
                    .RecordBytes,
                saved.Package.PersistObjects[savedVba.PersistId].RecordBytes);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));
        }

        [Fact]
        public void CapabilityContract_ReportsVbaProjectParity() {
            LegacyPptCapability capability = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.VbaProjects);

            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                capability.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.PptxToBinary);
            Assert.Contains("compressed or uncompressed", capability.Note);
            Assert.Contains("replacement, addition, and removal",
                capability.Note);
        }

        private static void SetVbaProject(PowerPointPresentation presentation,
            byte[] bytes) {
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart!;
            VbaProjectPart part = presentationPart.VbaProjectPart
                ?? presentationPart.AddNewPart<VbaProjectPart>();
            using var source = new MemoryStream(bytes, writable: false);
            part.FeedData(source);
        }

        private static byte[] ReadVbaProject(
            PowerPointPresentation presentation) {
            using Stream stream = presentation.OpenXmlDocument
                .PresentationPart!.VbaProjectPart!.GetStream(
                    FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream();
            stream.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] CreateVbaTestProject(string moduleName,
            string moduleText) {
            using var output = new MemoryStream();
            using (RootStorage root = RootStorage.Create(output,
                       CfbVersion.V3, StorageModeFlags.LeaveOpen)) {
                Storage vba = root.CreateStorage("VBA");
                using (CfbStream directory = vba.CreateStream("dir")) {
                    directory.Write(Array.Empty<byte>(), 0, 0);
                }
                using (CfbStream project = vba.CreateStream("_VBA_PROJECT")) {
                    project.Write(Array.Empty<byte>(), 0, 0);
                }
                using (CfbStream module = vba.CreateStream(moduleName)) {
                    byte[] moduleBytes = Encoding.UTF8.GetBytes(moduleText);
                    module.Write(moduleBytes, 0, moduleBytes.Length);
                }
                using (CfbStream project = root.CreateStream("PROJECT")) {
                    project.Write(Array.Empty<byte>(), 0, 0);
                }
            }
            return output.ToArray();
        }

        private static byte[] ConvertVbaStorageToCompressed(byte[] sourceBytes) {
            LegacyPptPresentation source = LegacyPptPresentation.Load(sourceBytes);
            LegacyPptVbaProject vba = Assert.IsType<LegacyPptVbaProject>(
                source.VbaProject);
            var persist = source.Package.PersistObjects[vba.PersistId];
            Assert.Equal(source.Package.PersistObjectOffsets.Values.Max(),
                persist.StreamOffset);

            byte[] compressed = CompressVbaZlib(vba.GetBytes());
            var payload = new byte[4 + compressed.Length];
            WriteVbaUInt32(payload, 0, checked((uint)vba.Length));
            Buffer.BlockCopy(compressed, 0, payload, 4, compressed.Length);
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
            int oldEditOffset = checked((int)source.Package.CurrentEditOffset);
            int editLength = checked(8 + (int)ReadVbaUInt32(
                source.Package.DocumentStream, oldEditOffset + 4));
            var edit = new byte[editLength];
            Buffer.BlockCopy(source.Package.DocumentStream, oldEditOffset,
                edit, 0, edit.Length);
            WriteVbaUInt32(edit, 20, directoryOffset);
            document.Write(edit, 0, edit.Length);

            byte[] currentUser = (byte[])source.Package.CurrentUserStream.Clone();
            WriteVbaUInt32(currentUser, 16, editOffset);
            return source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(
                    StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = document.ToArray(),
                    ["Current User"] = currentUser
                });
        }

        private static byte[] BuildVbaPersistDirectory(
            IReadOnlyDictionary<uint, uint> offsets) {
            KeyValuePair<uint, uint>[] entries = offsets.OrderBy(pair =>
                pair.Key).ToArray();
            using var payload = new MemoryStream();
            for (int index = 0; index < entries.Length;) {
                int count = 1;
                while (index + count < entries.Length && count < 0x0FFF
                       && entries[index + count].Key
                       == entries[index].Key + unchecked((uint)count)) {
                    count++;
                }
                WriteVbaUInt32(payload,
                    (unchecked((uint)count) << 20) | entries[index].Key);
                for (int item = 0; item < count; item++) {
                    WriteVbaUInt32(payload, entries[index + item].Value);
                }
                index += count;
            }
            return BuildVbaRecord(version: 0, instance: 0,
                type: 0x1772, payload.ToArray());
        }

        private static byte[] BuildVbaRecord(byte version, ushort instance,
            ushort type, byte[] payload) {
            var record = new byte[8 + payload.Length];
            WriteVbaUInt16(record, 0,
                unchecked((ushort)((instance << 4) | version)));
            WriteVbaUInt16(record, 2, type);
            WriteVbaUInt32(record, 4, checked((uint)payload.Length));
            Buffer.BlockCopy(payload, 0, record, 8, payload.Length);
            return record;
        }

        private static byte[] CompressVbaZlib(byte[] bytes) {
            using var output = new MemoryStream();
            output.WriteByte(0x78);
            output.WriteByte(0x9C);
            using (var deflate = new DeflateStream(output,
                       CompressionLevel.Optimal, leaveOpen: true)) {
                deflate.Write(bytes, 0, bytes.Length);
            }
            uint checksum = VbaAdler32(bytes);
            output.WriteByte((byte)(checksum >> 24));
            output.WriteByte((byte)(checksum >> 16));
            output.WriteByte((byte)(checksum >> 8));
            output.WriteByte((byte)checksum);
            return output.ToArray();
        }

        private static uint VbaAdler32(byte[] bytes) {
            const uint Modulus = 65521;
            uint a = 1;
            uint b = 0;
            foreach (byte value in bytes) {
                a = (a + value) % Modulus;
                b = (b + a) % Modulus;
            }
            return (b << 16) | a;
        }

        private static void WriteVbaUInt16(byte[] bytes, int offset,
            ushort value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteVbaUInt32(byte[] bytes, int offset,
            uint value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteVbaUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }

        private static uint ReadVbaUInt32(byte[] bytes, int offset) =>
            unchecked((uint)(bytes[offset]
                | (bytes[offset + 1] << 8)
                | (bytes[offset + 2] << 16)
                | (bytes[offset + 3] << 24)));
    }
}
