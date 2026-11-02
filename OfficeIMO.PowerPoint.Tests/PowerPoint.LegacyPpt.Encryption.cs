using OfficeIMO.Core.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.Tests.Pdf;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Theory]
        [InlineData(40)]
        [InlineData(56)]
        [InlineData(128)]
        public void LegacyEncryption_FreshBinaryRoundTripsTextAndPictures(
            int keySizeBits) {
            const string password = "Pässw0rd!";
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(23, 91, 177);
            byte[] encryptedBytes;
            using (PowerPointDocument source = PowerPointDocument.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.AddTextBox("Encrypted binary deck",global::OfficeIMO.OfficeLength.Emus(100000),global::OfficeIMO.OfficeLength.Emus(100000),                    global::OfficeIMO.OfficeLength.Emus(2200000),global::OfficeIMO.OfficeLength.Emus(600000));
                slide.AddImage(new MemoryStream(imageBytes),
                    OfficeIMO.Drawing.OfficeImageFormat.Png, global::OfficeIMO.OfficeLength.Emus(300000), global::OfficeIMO.OfficeLength.Emus(900000), global::OfficeIMO.OfficeLength.Emus(1800000), global::OfficeIMO.OfficeLength.Emus(1200000));
                encryptedBytes = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt, new PowerPointSaveOptions {
                        LegacyPptEncryptionKeySizeBits = keySizeBits
                    });
            }

            Assert.Throws<CryptographicException>(() =>
                LegacyPptDocument.Load(encryptedBytes));
            Assert.Throws<CryptographicException>(() =>
                LegacyPptDocument.Load(encryptedBytes,
                    new LegacyPptImportOptions { Password = "wrong" }));

            LegacyPptDocument legacy = LegacyPptDocument.Load(
                encryptedBytes, new LegacyPptImportOptions {
                    Password = password
                });
            Assert.True(legacy.WasEncryptedSource);
            Assert.Equal(keySizeBits, legacy.EncryptionKeySizeBits);
            Assert.True(legacy.EncryptedDocumentProperties);
            Assert.True(legacy.CreateImportReport().WasEncryptedSource);
            Assert.True(legacy.CreateImportReport()
                .EncryptedDocumentProperties);
            Assert.Contains(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "ppt-encryption-decrypted");
            Assert.Contains(legacy.Slides[0].Shapes,
                shape => shape.Text == "Encrypted binary deck");
            Assert.Equal(imageBytes, Assert.Single(legacy.BlipStoreEntries)
                .ImageBytes);

            using var input = new MemoryStream(encryptedBytes);
            using PowerPointDocument projected =
                PowerPointDocument.LoadEncrypted(input, password);
            Assert.Equal(PowerPointFileFormat.Ppt, projected.SourceFormat);
            Assert.Contains(projected.Slides[0].TextBoxes,
                textBox => textBox.Text == "Encrypted binary deck");
            Assert.Single(projected.Slides[0].Pictures);
        }

        [Fact]
        public void LegacyEncryption_RejectsOverlappingPersistRangesBeforeDecrypting() {
            const string password = "overlap-pass";
            byte[] encryptedBytes = CreateEncryptedLegacyPresentation(
                password);
            Assert.True(OfficeCompoundFileReader.TryRead(encryptedBytes,
                out OfficeCompoundFile? compound, out string? readError),
                readError);
            OfficeCompoundFile sourceCompound = Assert.IsType<
                OfficeCompoundFile>(compound);
            byte[] document = (byte[])sourceCompound.Streams[
                "PowerPoint Document"].Clone();
            LegacyPptCurrentUserAtom currentUser =
                LegacyPptCurrentUserAtom.Read(sourceCompound.Streams[
                    "Current User"]);
            LegacyPptRecord edit = LegacyPptRecordReader.ReadSingle(
                document, checked((int)currentUser.CurrentEditOffset),
                new LegacyPptImportOptions());
            LegacyPptRecord directory = LegacyPptRecordReader.ReadSingle(
                document, checked((int)edit.ReadUInt32(12)),
                new LegacyPptImportOptions());
            var entries = new List<(uint PersistId, int ValueOffset,
                uint StreamOffset)>();
            int position = 0;
            while (position < directory.PayloadLength) {
                uint packed = directory.ReadUInt32(position);
                position += 4;
                uint firstId = packed & 0x000FFFFF;
                int count = unchecked((int)(packed >> 20));
                for (int index = 0; index < count; index++) {
                    entries.Add((checked(firstId + unchecked((uint)index)),
                        checked(directory.PayloadOffset + position),
                        directory.ReadUInt32(position)));
                    position += 4;
                }
            }

            uint encryptionPersistId = edit.ReadUInt32(28);
            uint encryptionOffset = entries.Single(entry =>
                entry.PersistId == encryptionPersistId).StreamOffset;
            int encryptionRecordOffset = checked((int)encryptionOffset);
            int encryptionRecordLength = checked(8 + (int)
                ReadCompoundUInt32(document, encryptionRecordOffset + 4));
            var encryptionHeader = new byte[encryptionRecordLength - 8];
            Buffer.BlockCopy(document, encryptionRecordOffset + 8,
                encryptionHeader, 0, encryptionHeader.Length);
            OfficeBinaryRc4CryptoApiSession session =
                OfficeBinaryRc4CryptoApiSession.Open(encryptionHeader,
                    password);

            var inspected = new List<(uint PersistId, int ValueOffset,
                uint StreamOffset, int RecordLength)>();
            foreach ((uint persistId, int valueOffset, uint streamOffset)
                     in entries.Where(entry => entry.PersistId !=
                         encryptionPersistId)) {
                int streamPosition = checked((int)streamOffset);
                var header = new byte[8];
                Buffer.BlockCopy(document, streamPosition, header, 0,
                    header.Length);
                session.TransformInPlace(header, 0, header.Length,
                    persistId);
                int recordLength = checked(8 + (int)
                    ReadCompoundUInt32(header, 4));
                inspected.Add((persistId, valueOffset, streamOffset,
                    recordLength));
            }
            (uint PersistId, int ValueOffset, uint StreamOffset,
                int RecordLength) owner = inspected.First(entry =>
                    entry.RecordLength >= 16);
            (uint PersistId, int ValueOffset, uint StreamOffset,
                int RecordLength) alias = inspected.First(entry =>
                    entry.PersistId != owner.PersistId);
            uint overlappingOffset = checked(owner.StreamOffset + 8U);
            var overlappingHeader = new byte[8];
            overlappingHeader[3] = 0x10;
            session.TransformInPlace(overlappingHeader, 0,
                overlappingHeader.Length, alias.PersistId);
            Buffer.BlockCopy(overlappingHeader, 0, document,
                checked((int)overlappingOffset), overlappingHeader.Length);
            WriteUInt32(document, alias.ValueOffset, overlappingOffset);

            byte[] maliciousBytes = OfficeCompoundFileWriter.Rewrite(
                sourceCompound, new Dictionary<string, byte[]> {
                    ["PowerPoint Document"] = document
                });
            Assert.True(OfficeCompoundFileReader.TryRead(maliciousBytes,
                out OfficeCompoundFile? maliciousCompound,
                out string? maliciousError), maliciousError);
            OfficeCompoundFile maliciousSource = Assert.IsType<
                OfficeCompoundFile>(maliciousCompound);
            var options = new LegacyPptImportOptions {
                Password = password
            };

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptRc4CryptoApi
                    .DecryptPackage(maliciousBytes, maliciousSource,
                        options, new LegacyPptRecordTraversalBudget(
                            options.MaxRecordCount),
                        new LegacyPptDecodedStorageBudget(
                            options.MaxDecodedStorageBytes),
                        CancellationToken.None, out _, out _));

            Assert.Contains("overlap", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void LegacyEncryption_ChargesDecryptedPersistObjectsToStorageBudget() {
            const string password = "persist-budget-pass";
            byte[] encryptedBytes = CreateEncryptedLegacyPresentation(
                password);

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptDocument.Load(
                    encryptedBytes, new LegacyPptImportOptions {
                        Password = password,
                        MaxDecodedStorageBytes = 1
                    }));

            Assert.Contains("aggregate decoded embedded-storage",
                exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void LegacyEncryption_NormalLoadAcceptsPasswordInBinaryOptions() {
            const string password = "normal-flow";
            byte[] encryptedBytes = CreateEncryptedLegacyPresentation(password);
            var options = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    Password = password
                }
            };

            using var input = new MemoryStream(encryptedBytes);
            using PowerPointDocument presentation =
                PowerPointDocument.Load(input, options);

            Assert.Equal(PowerPointFileFormat.Ppt, presentation.SourceFormat);
            Assert.Contains(presentation.LegacyPptImportDiagnostics,
                diagnostic => diagnostic.Code == "ppt-encryption-decrypted");
            Assert.Contains(presentation.Slides[0].TextBoxes,
                textBox => textBox.Text == "Original secret");
        }

        [Fact]
        public async Task LegacyEncryption_LoadEncryptedRejectsPlainBinaryInput() {
            byte[] plainBytes;
            using (PowerPointDocument source =
                   PowerPointDocument.Create()) {
                source.AddSlide().AddTextBox("Plain binary deck");
                plainBytes = source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
            }

            using (var stream = new MemoryStream(plainBytes)) {
                InvalidDataException failure = Assert.Throws<InvalidDataException>(
                    () => PowerPointDocument.LoadEncrypted(stream,
                        "ignored-password"));
                Assert.Contains("not password-encrypted", failure.Message);
            }
            using (var stream = new MemoryStream(plainBytes)) {
                InvalidDataException failure = await Assert.ThrowsAsync<
                    InvalidDataException>(() => PowerPointDocument
                    .LoadEncryptedAsync(stream, "ignored-password"));
                Assert.Contains("not password-encrypted", failure.Message);
            }

            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".ppt");
            try {
                File.WriteAllBytes(path, plainBytes);
                InvalidDataException pathFailure = Assert.Throws<
                    InvalidDataException>(() => PowerPointDocument
                    .LoadEncrypted(path, "ignored-password"));
                Assert.Contains("not password-encrypted",
                    pathFailure.Message);
                InvalidDataException asyncPathFailure = await Assert
                    .ThrowsAsync<InvalidDataException>(() =>
                        PowerPointDocument.LoadEncryptedAsync(path,
                            "ignored-password"));
                Assert.Contains("not password-encrypted",
                    asyncPathFailure.Message);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void LegacyEncryption_NormalPathSaveRetainsSourceEncryption() {
            const string password = "normal-save-pass";
            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".ppt");
            try {
                using (PowerPointDocument source =
                       PowerPointDocument.Create()) {
                    source.AddSlide().AddTextBox("Before encrypted edit");
                    source.SaveEncrypted(path, password,
                        new PowerPointSaveOptions {
                            LegacyPptEncryptionKeySizeBits = 56,
                            LegacyPptEncryptDocumentProperties = false
                        });
                }

                using (PowerPointDocument loaded =
                       PowerPointDocument.Load(path,
                           new PowerPointLoadOptions {
                               LegacyPptImportOptions =
                                   new LegacyPptImportOptions {
                                       Password = password
                                   }
                           })) {
                    loaded.Slides[0].TextBoxes.Single(textBox =>
                        textBox.Text == "Before encrypted edit").Text =
                            "After encrypted edit";
                    loaded.Save();
                }

                byte[] savedBytes = File.ReadAllBytes(path);
                Assert.Throws<CryptographicException>(() =>
                    LegacyPptDocument.Load(savedBytes));
                LegacyPptDocument reopened =
                    LegacyPptDocument.Load(savedBytes,
                        new LegacyPptImportOptions { Password = password });
                Assert.Equal(56, reopened.EncryptionKeySizeBits);
                Assert.False(reopened.EncryptedDocumentProperties);
                Assert.Contains(reopened.Slides[0].Shapes,
                    shape => shape.Text == "After encrypted edit");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void LegacyEncryption_ExactNoOpBinarySavePreservesCiphertext() {
            const string password = "exact-pass";
            byte[] encryptedBytes = CreateEncryptedLegacyPresentation(
                password);
            using var input = new MemoryStream(encryptedBytes);
            using PowerPointDocument presentation =
                PowerPointDocument.Load(input,
                    new PowerPointLoadOptions {
                        LegacyPptImportOptions =
                            new LegacyPptImportOptions {
                                Password = password
                            }
                    });

            Assert.Equal(encryptedBytes,
                presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }

        [Fact]
        public void LegacyEncryption_RespectsLegacySignatureMutationPolicy() {
            byte[] signedBytes = CreateLegacySignatureFixture();
            using var input = new MemoryStream(signedBytes);
            using PowerPointDocument presentation =
                PowerPointDocument.Load(input);

            Assert.Throws<PowerPointSignedDocumentMutationException>(() =>
                presentation.ToEncryptedBytes("signed-pass",
                    PowerPointFileFormat.Ppt));

            presentation.SignatureMutationPolicy =
                OfficeSignatureMutationPolicy
                    .RemoveInvalidatedSignatures;
            byte[] encrypted = presentation.ToEncryptedBytes("signed-pass",
                PowerPointFileFormat.Ppt);
            LegacyPptDocument reopened =
                LegacyPptDocument.Load(encrypted,
                    new LegacyPptImportOptions {
                        Password = "signed-pass"
                    });
            Assert.DoesNotContain(
                reopened.Package.CopyCompoundStreams().Keys,
                path => path.Equals("_signatures",
                            StringComparison.OrdinalIgnoreCase)
                        || path.Equals("_xmlsignatures",
                            StringComparison.OrdinalIgnoreCase)
                        || path.StartsWith("_xmlsignatures/",
                            StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void LegacyEncryption_DefaultEncryptsAndRestoresDocumentProperties() {
            const string password = "property-pass";
            const string title = "Confidential binary metadata";
            byte[] encryptedBytes;
            using (PowerPointDocument source =
                   PowerPointDocument.Create()) {
                source.AddSlide().AddTextBox("Visible slide text");
                source.BuiltinDocumentProperties.Title = title;
                encryptedBytes = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt);
            }

            Assert.False(ContainsBytes(encryptedBytes,
                Encoding.Unicode.GetBytes(title)));

            using var input = new MemoryStream(encryptedBytes);
            using PowerPointDocument reopened =
                PowerPointDocument.LoadEncrypted(input, password);
            Assert.Equal(title, reopened.BuiltinDocumentProperties.Title);
        }

        [Fact]
        public void LegacyEncryption_CanLeaveDocumentPropertiesInClearText() {
            const string password = "clear-property-pass";
            const string title = "Public binary metadata";
            byte[] encryptedBytes;
            using (PowerPointDocument source =
                   PowerPointDocument.Create()) {
                source.AddSlide().AddTextBox("Encrypted slide text");
                source.BuiltinDocumentProperties.Title = title;
                encryptedBytes = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LegacyPptEncryptDocumentProperties = false
                    });
            }

            Assert.True(ContainsBytes(encryptedBytes,
                Encoding.Unicode.GetBytes(title)));

            using var input = new MemoryStream(encryptedBytes);
            using PowerPointDocument reopened =
                PowerPointDocument.LoadEncrypted(input, password);
            Assert.Equal(title, reopened.BuiltinDocumentProperties.Title);
        }

        [Fact]
        public void LegacyEncryption_OpaqueStreamsRequireExplicitClearTextOptIn() {
            const string sourcePassword = "source-pass";
            const string targetPassword = "target-pass";
            byte[] plainBytes;
            using (PowerPointDocument source = PowerPointDocument.Create()) {
                source.AddSlide().AddTextBox("Original secret",global::OfficeIMO.OfficeLength.Emus(100000),                    global::OfficeIMO.OfficeLength.Emus(100000),global::OfficeIMO.OfficeLength.Emus(2200000),global::OfficeIMO.OfficeLength.Emus(600000));
                plainBytes = source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
            }
            LegacyPptDocument plain = LegacyPptDocument.Load(plainBytes);
            byte[] opaqueBytes = "opaque-vendor-stream"u8.ToArray();
            byte[] withOpaque = plain.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]> {
                    ["VendorData/opaque"] = opaqueBytes
                });
            NotSupportedException sourceFailure = Assert.Throws<
                NotSupportedException>(() => LegacyPptRc4CryptoApi
                .EncryptPackage(withOpaque, sourcePassword));
            Assert.Contains("VendorData/opaque", sourceFailure.Message,
                StringComparison.Ordinal);
            byte[] encryptedSource = LegacyPptRc4CryptoApi.EncryptPackage(
                withOpaque, sourcePassword,
                allowUnencryptedCompoundStreams: true);

            byte[] encryptedEdited;
            using (var input = new MemoryStream(encryptedSource))
            using (PowerPointDocument presentation =
                   PowerPointDocument.LoadEncrypted(input,
                       sourcePassword)) {
                Assert.Single(presentation.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Original secret").Text =
                    "Edited secret";
                Assert.Throws<NotSupportedException>(() => presentation
                    .ToEncryptedBytes(targetPassword,
                        PowerPointFileFormat.Ppt));
                encryptedEdited = presentation.ToEncryptedBytes(
                    targetPassword, PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LegacyPptAllowUnencryptedCompoundStreams = true
                    });
            }

            Assert.Throws<CryptographicException>(() =>
                LegacyPptDocument.Load(encryptedEdited,
                    new LegacyPptImportOptions { Password = sourcePassword }));
            LegacyPptDocument edited = LegacyPptDocument.Load(
                encryptedEdited, new LegacyPptImportOptions {
                    Password = targetPassword
                });
            Assert.Contains(edited.Slides[0].Shapes,
                shape => shape.Text == "Edited secret");
            Assert.Equal(opaqueBytes,
                edited.Package.CopyCompoundStreams()["VendorData/opaque"]);
            Assert.Single(edited.Package.UserEdits);
        }

        [Theory]
        [InlineData(".ppt", PowerPointFileFormat.Ppt)]
        [InlineData(".pot", PowerPointFileFormat.Pot)]
        [InlineData(".pps", PowerPointFileFormat.Pps)]
        public void LegacyEncryption_SaveEncryptedPathRoutesBinaryVariants(
            string extension, PowerPointFileFormat expectedFormat) {
            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + extension);
            try {
                using (PowerPointDocument source =
                       PowerPointDocument.Create()) {
                    source.AddSlide().AddTextBox("Encrypted variant");
                    source.SaveEncrypted(path, "variant-pass");
                }

                using PowerPointDocument loaded =
                    PowerPointDocument.LoadEncrypted(path,
                        "variant-pass");
                Assert.Equal(expectedFormat, loaded.SourceFormat);
                Assert.Null(loaded.SourcePath);
                Assert.Contains(loaded.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Encrypted variant");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task LegacyEncryption_AsyncPathAndStreamLoadsRouteBinary() {
            const string password = "async-pass";
            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pps");
            try {
                using (PowerPointDocument source =
                       PowerPointDocument.Create()) {
                    source.AddSlide().AddTextBox("Async encrypted binary");
                    source.SaveEncrypted(path, password);
                }

                using PowerPointDocument fromPath =
                    await PowerPointDocument.LoadEncryptedAsync(path,
                        password);
                Assert.Equal(PowerPointFileFormat.Pps,
                    fromPath.SourceFormat);

                byte[] bytes = File.ReadAllBytes(path);
                using var encryptedStream = new MemoryStream(bytes);
                using PowerPointDocument fromStream =
                    await PowerPointDocument.LoadEncryptedAsync(
                        encryptedStream, password);
                Assert.Equal(PowerPointFileFormat.Ppt,
                    fromStream.SourceFormat);
                Assert.Contains(fromStream.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Async encrypted binary");

                using var normalStream = new MemoryStream(bytes);
                using PowerPointDocument normalFlow =
                    await PowerPointDocument.LoadAsync(normalStream,
                        new PowerPointLoadOptions {
                            LegacyPptImportOptions =
                                new LegacyPptImportOptions {
                                    Password = password
                                }
                        });
                Assert.Contains(normalFlow.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Async encrypted binary");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Theory]
        [InlineData(PowerPointFileFormat.Ppt)]
        [InlineData(PowerPointFileFormat.Pptx)]
        public async Task Encryption_AsyncLoadCancelsAfterInputBuffering(
            PowerPointFileFormat format) {
            const string password = "cancel-after-buffering";
            byte[] encrypted;
            using (PowerPointDocument source =
                   PowerPointDocument.Create()) {
                source.AddSlide().AddTextBox("Cancellation boundary");
                encrypted = source.ToEncryptedBytes(password, format);
            }

            using var cancellation = new CancellationTokenSource();
            using var input = new CancelAtEndAsyncReadStream(encrypted,
                cancellation.Cancel);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                PowerPointDocument.LoadEncryptedAsync(input, password,
                    cancellationToken: cancellation.Token));
        }

        [Fact]
        public void LegacyEncryption_StreamSaveRetainsImportedBinaryFormat() {
            const string sourcePassword = "stream-source";
            const string targetPassword = "stream-target";
            byte[] encrypted = CreateEncryptedLegacyPresentation(
                sourcePassword);

            using var input = new MemoryStream(encrypted);
            using PowerPointDocument presentation =
                PowerPointDocument.LoadEncrypted(input, sourcePassword);
            Assert.Equal(PowerPointFileFormat.Ppt,
                presentation.SourceFormat);
            presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "Original secret").Text =
                    "Re-encrypted stream";

            using var output = new MemoryStream();
            presentation.SaveEncrypted(output, targetPassword);
            using var reopenedInput = new MemoryStream(output.ToArray());
            using PowerPointDocument reopened =
                PowerPointDocument.LoadEncrypted(reopenedInput,
                    targetPassword);
            Assert.Contains(reopened.Slides[0].TextBoxes,
                textBox => textBox.Text == "Re-encrypted stream");
        }

        [Theory]
        [InlineData(32)]
        [InlineData(41)]
        [InlineData(136)]
        public void LegacyEncryption_RejectsInvalidRc4KeySizes(int keySizeBits) {
            using PowerPointDocument source =
                PowerPointDocument.Create();
            source.AddSlide().AddTextBox("Invalid key size");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                source.ToEncryptedBytes("key-pass", PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LegacyPptEncryptionKeySizeBits = keySizeBits
                    }));
        }

        private static byte[] CreateEncryptedLegacyPresentation(
            string password) {
            using PowerPointDocument source =
                PowerPointDocument.Create();
            source.AddSlide().AddTextBox("Original secret",global::OfficeIMO.OfficeLength.Emus(100000),                global::OfficeIMO.OfficeLength.Emus(100000),global::OfficeIMO.OfficeLength.Emus(2200000),global::OfficeIMO.OfficeLength.Emus(600000));
            return source.ToEncryptedBytes(password, PowerPointFileFormat.Ppt);
        }

        private static bool ContainsBytes(byte[] source, byte[] value) {
            if (value.Length == 0) return true;
            for (int offset = 0; offset <= source.Length - value.Length;
                 offset++) {
                int index = 0;
                while (index < value.Length
                       && source[offset + index] == value[index]) {
                    index++;
                }
                if (index == value.Length) return true;
            }
            return false;
        }

        private sealed class CancelAtEndAsyncReadStream : Stream {
            private readonly MemoryStream _inner;
            private readonly Action _cancel;

            public CancelAtEndAsyncReadStream(byte[] bytes, Action cancel) {
                _inner = new MemoryStream(bytes, writable: false);
                _cancel = cancel;
            }

            public override bool CanRead => true;
            public override bool CanSeek => true;
            public override bool CanWrite => false;
            public override long Length => _inner.Length;
            public override long Position {
                get => _inner.Position;
                set => _inner.Position = value;
            }

            public override int Read(byte[] buffer, int offset, int count) =>
                _inner.Read(buffer, offset, count);

            public override async Task<int> ReadAsync(byte[] buffer,
                int offset, int count,
                CancellationToken cancellationToken) {
                int read = await _inner.ReadAsync(buffer, offset, count,
                    CancellationToken.None).ConfigureAwait(false);
                if (read == 0) _cancel();
                return read;
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) =>
                _inner.Seek(offset, origin);
            public override void SetLength(long value) =>
                throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset,
                int count) => throw new NotSupportedException();

            protected override void Dispose(bool disposing) {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
