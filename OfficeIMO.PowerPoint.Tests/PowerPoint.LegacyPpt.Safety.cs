using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptSafetyTests {
        [Fact]
        public void RecordReader_RejectsOversizedDeclaredPayloadWithoutAllocation() {
            byte[] record = CreateRecord(version: 0, payload: Array.Empty<byte>());
            WriteUInt32(record, 4, uint.MaxValue);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                LegacyPptRecordReader.ReadSingle(record, 0,
                    new LegacyPptImportOptions()));

            Assert.Contains("too large", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void RecordReader_EnforcesNestingDepthBudget() {
            byte[] atom = CreateRecord(version: 0, payload: Array.Empty<byte>());
            byte[] nested = CreateRecord(version: 0x0F,
                payload: CreateRecord(version: 0x0F, payload: atom));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                LegacyPptRecordReader.ReadSingle(nested, 0,
                    new LegacyPptImportOptions { MaxRecordDepth = 1 }));

            Assert.Contains("nesting depth", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PackageReader_EnforcesFileWideRecordCountBudget() {
            byte[] bytes = CreatePresentationBytes();
            LegacyPptPresentation source = LegacyPptPresentation.Load(bytes);
            var unrestricted = new LegacyPptImportOptions();
            int maximumSingleTree = source.Package.PersistObjects.Values
                .Select(persistObject => LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0, unrestricted)
                    .DescendantsAndSelf().Count())
                .Max();
            int combinedPersistTreeCount = source.Package.PersistObjects.Values
                .Sum(persistObject => LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0, unrestricted)
                    .DescendantsAndSelf().Count());

            Assert.True(combinedPersistTreeCount > maximumSingleTree);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                LegacyPptPresentation.Load(bytes, new LegacyPptImportOptions {
                    MaxRecordCount = maximumSingleTree
                }));

            Assert.Contains("record count", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PackageReader_EnforcesDocumentStreamSizeBudget() {
            byte[] bytes = CreatePresentationBytes();
            LegacyPptPresentation source = LegacyPptPresentation.Load(bytes);
            int limit = source.Package.DocumentStream.Length - 1;

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                LegacyPptPresentation.Load(bytes,
                    new LegacyPptImportOptions { MaxInputBytes = limit }));

            Assert.Contains("exceeds", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PackageReader_RejectsCyclicUserEditChain() {
            byte[] bytes = CreatePresentationBytes();
            LegacyPptPresentation source = LegacyPptPresentation.Load(bytes);
            byte[] document = (byte[])source.Package.DocumentStream.Clone();
            int previousEditOffset = checked((int)source.Package.CurrentEditOffset + 16);
            WriteUInt32(document, previousEditOffset,
                source.Package.CurrentEditOffset);
            byte[] cyclic = source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = document
                });

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                LegacyPptPresentation.Load(cyclic));

            Assert.Contains("cycle", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        private static byte[] CreatePresentationBytes() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Safety fixture");
            presentation.AddSlide().AddTextBox("Second safety slide");
            return presentation.ToBytes(PowerPointFileFormat.Ppt);
        }

        private static byte[] CreateRecord(byte version, byte[] payload) {
            byte[] record = new byte[checked(8 + payload.Length)];
            WriteUInt16(record, 0, version);
            WriteUInt16(record, 2, 0x1000);
            WriteUInt32(record, 4, checked((uint)payload.Length));
            Buffer.BlockCopy(payload, 0, record, 8, payload.Length);
            return record;
        }

        private static void WriteUInt16(byte[] target, int offset, ushort value) {
            target[offset] = unchecked((byte)value);
            target[offset + 1] = unchecked((byte)(value >> 8));
        }

        private static void WriteUInt32(byte[] target, int offset, uint value) {
            target[offset] = unchecked((byte)value);
            target[offset + 1] = unchecked((byte)(value >> 8));
            target[offset + 2] = unchecked((byte)(value >> 16));
            target[offset + 3] = unchecked((byte)(value >> 24));
        }
    }
}
