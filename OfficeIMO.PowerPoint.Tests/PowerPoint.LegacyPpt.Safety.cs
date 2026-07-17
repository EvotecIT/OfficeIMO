using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
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
        public void PackageReader_EnforcesInputBudgetBeforeBufferingPathOrStream() {
            byte[] bytes = CreatePresentationBytes();
            int limit = bytes.Length - 1;
            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid().ToString("N") + ".ppt");
            try {
                File.WriteAllBytes(path, bytes);
                var options = new LegacyPptImportOptions {
                    MaxInputBytes = limit
                };

                InvalidDataException pathException = Assert.Throws<
                    InvalidDataException>(() => LegacyPptPresentation.Load(
                    path, options));
                using var stream = new ReadGuardStream(bytes.Length);
                InvalidDataException streamException = Assert.Throws<
                    InvalidDataException>(() => LegacyPptPresentation.Load(
                    stream, options));

                Assert.Contains("exceeds", pathException.Message,
                    StringComparison.OrdinalIgnoreCase);
                Assert.Contains("exceeds", streamException.Message,
                    StringComparison.OrdinalIgnoreCase);
                Assert.Equal(0, stream.ReadCount);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task PresentationFacade_EnforcesInputBudgetBeforeBuffering() {
            const int length = 4096;
            var loadOptions = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    MaxInputBytes = length - 1
                }
            };

            using var loadStream = new ReadGuardStream(length);
            using var encryptedStream = new ReadGuardStream(length);
            using var loadAsyncStream = new ReadGuardStream(length);
            using var encryptedAsyncStream = new ReadGuardStream(length);

            Assert.Throws<InvalidDataException>(() =>
                PowerPointPresentation.Load(loadStream, loadOptions));
            Assert.Throws<InvalidDataException>(() =>
                PowerPointPresentation.LoadEncrypted(encryptedStream,
                    "password", loadOptions));
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                PowerPointPresentation.LoadAsync(loadAsyncStream,
                    loadOptions));
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                PowerPointPresentation.LoadEncryptedAsync(
                    encryptedAsyncStream, "password", loadOptions));

            Assert.Equal(0, loadStream.ReadCount);
            Assert.Equal(0, encryptedStream.ReadCount);
            Assert.Equal(0, loadAsyncStream.ReadCount);
            Assert.Equal(0, encryptedAsyncStream.ReadCount);
        }

        [Fact]
        public void PackageReader_EnforcesImportWideDecodedStorageBudget() {
            const string Password = "storage-budget";
            byte[] firstStorage = CreateOleTestStorage("First storage");
            byte[] secondStorage = CreateOleTestStorage("Second storage");
            byte[] bytes;
            byte[] encryptedBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide();
                using var first = new MemoryStream(firstStorage,
                    writable: false);
                using var second = new MemoryStream(secondStorage,
                    writable: false);
                slide.AddOleObject(first, "Package");
                slide.AddOleObject(second, "Package");
                bytes = created.ToBytes(PowerPointFileFormat.Ppt);
                encryptedBytes = created.ToEncryptedBytes(Password,
                    PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation unrestricted =
                LegacyPptPresentation.Load(bytes);
            Assert.Equal(2, unrestricted.OleObjects.Count);
            long decodedBytes = unrestricted.OleObjects.Sum(item =>
                (long)item.Length);

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptPresentation.Load(bytes,
                new LegacyPptImportOptions {
                    MaxDecodedStorageBytes = decodedBytes - 1
                }));

            Assert.Contains("aggregate decoded embedded-storage",
                exception.Message, StringComparison.OrdinalIgnoreCase);

            var loadOptions = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    MaxDecodedStorageBytes = decodedBytes - 1
                }
            };
            using var input = new MemoryStream(bytes, writable: false);
            Assert.Throws<InvalidDataException>(() =>
                PowerPointPresentation.Load(input, loadOptions));

            using var encryptedInput = new MemoryStream(encryptedBytes,
                writable: false);
            Assert.Throws<InvalidDataException>(() =>
                PowerPointPresentation.LoadEncrypted(encryptedInput,
                    Password, loadOptions));
        }

        [Fact]
        public void OleStorageDecoder_ReservesCompressedExpansionBeforeDecode() {
            byte[] decoded = CreateOleTestStorage("Compressed expansion");
            byte[] compressed = CompressVbaZlib(decoded);
            var payload = new byte[checked(4 + compressed.Length)];
            WriteVbaUInt32(payload, 0, checked((uint)decoded.Length));
            Buffer.BlockCopy(compressed, 0, payload, 4,
                compressed.Length);
            byte[] record = BuildVbaRecord(version: 0, instance: 1,
                type: 0x1011, payload);
            var persistObject = new LegacyPptPersistObject(1, 0, 0x1011,
                record);
            var options = new LegacyPptImportOptions {
                MaxDecodedStorageBytes = decoded.Length - 1
            };
            var recordBudget = new LegacyPptRecordTraversalBudget(
                options.MaxRecordCount);
            var decodedBudget = new LegacyPptDecodedStorageBudget(
                options.MaxDecodedStorageBytes);

            Assert.False(LegacyPptOleStorageCodec.TryDecode(persistObject,
                options, recordBudget, decodedBudget,
                out byte[] decodedStorage, out bool wasCompressed,
                out string? reason));

            Assert.Empty(decodedStorage);
            Assert.True(wasCompressed);
            Assert.Equal(0, decodedBudget.DecodedBytes);
            Assert.Contains("aggregate decoded embedded-storage", reason,
                StringComparison.OrdinalIgnoreCase);
            Assert.Throws<InvalidDataException>(() =>
                decodedBudget.ThrowIfExceeded());
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

        private sealed class ReadGuardStream : Stream {
            private long _position;

            public ReadGuardStream(long length) {
                Length = length;
            }

            public int ReadCount { get; private set; }
            public override bool CanRead => true;
            public override bool CanSeek => true;
            public override bool CanWrite => false;
            public override long Length { get; }
            public override long Position {
                get => _position;
                set => _position = value;
            }

            public override int Read(byte[] buffer, int offset, int count) {
                ReadCount++;
                throw new InvalidOperationException(
                    "The length guard should reject the stream before reading.");
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) {
                _position = origin switch {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => checked(_position + offset),
                    SeekOrigin.End => checked(Length + offset),
                    _ => throw new ArgumentOutOfRangeException(nameof(origin))
                };
                return _position;
            }
            public override void SetLength(long value) =>
                throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset,
                int count) => throw new NotSupportedException();
        }
    }
}
