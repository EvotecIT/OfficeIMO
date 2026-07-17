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
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.AddTextBox("Encrypted binary deck", 100000, 100000,
                    2200000, 600000);
                slide.AddPicture(new MemoryStream(imageBytes),
                    ImagePartType.Png, 300000, 900000, 1800000, 1200000);
                encryptedBytes = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt, new PowerPointSaveOptions {
                        LegacyPptEncryptionKeySizeBits = keySizeBits
                    });
            }

            Assert.Throws<CryptographicException>(() =>
                LegacyPptPresentation.Load(encryptedBytes));
            Assert.Throws<CryptographicException>(() =>
                LegacyPptPresentation.Load(encryptedBytes,
                    new LegacyPptImportOptions { Password = "wrong" }));

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(
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
                diagnostic.Code == "PPT-ENCRYPTION-DECRYPTED");
            Assert.Contains(legacy.Slides[0].Shapes,
                shape => shape.Text == "Encrypted binary deck");
            Assert.Equal(imageBytes, Assert.Single(legacy.BlipStoreEntries)
                .ImageBytes);

            using var input = new MemoryStream(encryptedBytes);
            using PowerPointPresentation projected =
                PowerPointPresentation.LoadEncrypted(input, password);
            Assert.Equal(PowerPointFileFormat.Ppt, projected.SourceFormat);
            Assert.Contains(projected.Slides[0].TextBoxes,
                textBox => textBox.Text == "Encrypted binary deck");
            Assert.Single(projected.Slides[0].Pictures);
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
            using PowerPointPresentation presentation =
                PowerPointPresentation.Load(input, options);

            Assert.Equal(PowerPointFileFormat.Ppt, presentation.SourceFormat);
            Assert.Contains(presentation.LegacyPptImportDiagnostics,
                diagnostic => diagnostic.Code == "PPT-ENCRYPTION-DECRYPTED");
            Assert.Contains(presentation.Slides[0].TextBoxes,
                textBox => textBox.Text == "Original secret");
        }

        [Fact]
        public async Task LegacyEncryption_LoadEncryptedRejectsPlainBinaryInput() {
            byte[] plainBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Plain binary deck");
                plainBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            using (var stream = new MemoryStream(plainBytes)) {
                InvalidDataException failure = Assert.Throws<InvalidDataException>(
                    () => PowerPointPresentation.LoadEncrypted(stream,
                        "ignored-password"));
                Assert.Contains("not password-encrypted", failure.Message);
            }
            using (var stream = new MemoryStream(plainBytes)) {
                InvalidDataException failure = await Assert.ThrowsAsync<
                    InvalidDataException>(() => PowerPointPresentation
                    .LoadEncryptedAsync(stream, "ignored-password"));
                Assert.Contains("not password-encrypted", failure.Message);
            }

            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".ppt");
            try {
                File.WriteAllBytes(path, plainBytes);
                InvalidDataException pathFailure = Assert.Throws<
                    InvalidDataException>(() => PowerPointPresentation
                    .LoadEncrypted(path, "ignored-password"));
                Assert.Contains("not password-encrypted",
                    pathFailure.Message);
                InvalidDataException asyncPathFailure = await Assert
                    .ThrowsAsync<InvalidDataException>(() =>
                        PowerPointPresentation.LoadEncryptedAsync(path,
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
                using (PowerPointPresentation source =
                       PowerPointPresentation.Create()) {
                    source.AddSlide().AddTextBox("Before encrypted edit");
                    source.SaveEncrypted(path, password,
                        new PowerPointSaveOptions {
                            LegacyPptEncryptionKeySizeBits = 56,
                            LegacyPptEncryptDocumentProperties = false
                        });
                }

                using (PowerPointPresentation loaded =
                       PowerPointPresentation.Load(path,
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
                    LegacyPptPresentation.Load(savedBytes));
                LegacyPptPresentation reopened =
                    LegacyPptPresentation.Load(savedBytes,
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
            using PowerPointPresentation presentation =
                PowerPointPresentation.Load(input,
                    new PowerPointLoadOptions {
                        LegacyPptImportOptions =
                            new LegacyPptImportOptions {
                                Password = password
                            }
                    });

            Assert.Equal(encryptedBytes,
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void LegacyEncryption_RespectsLegacySignatureMutationPolicy() {
            byte[] signedBytes = CreateLegacySignatureFixture();
            using var input = new MemoryStream(signedBytes);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Load(input);

            Assert.Throws<PowerPointSignedPresentationMutationException>(() =>
                presentation.ToEncryptedBytes("signed-pass",
                    PowerPointFileFormat.Ppt));

            presentation.SignatureMutationPolicy =
                PowerPointSignatureMutationPolicy
                    .RemoveInvalidatedSignatures;
            byte[] encrypted = presentation.ToEncryptedBytes("signed-pass",
                PowerPointFileFormat.Ppt);
            LegacyPptPresentation reopened =
                LegacyPptPresentation.Load(encrypted,
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
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Visible slide text");
                source.BuiltinDocumentProperties.Title = title;
                encryptedBytes = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt);
            }

            Assert.False(ContainsBytes(encryptedBytes,
                Encoding.Unicode.GetBytes(title)));

            using var input = new MemoryStream(encryptedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.LoadEncrypted(input, password);
            Assert.Equal(title, reopened.BuiltinDocumentProperties.Title);
        }

        [Fact]
        public void LegacyEncryption_CanLeaveDocumentPropertiesInClearText() {
            const string password = "clear-property-pass";
            const string title = "Public binary metadata";
            byte[] encryptedBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
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
            using PowerPointPresentation reopened =
                PowerPointPresentation.LoadEncrypted(input, password);
            Assert.Equal(title, reopened.BuiltinDocumentProperties.Title);
        }

        [Fact]
        public void LegacyEncryption_EditedImportedDeckCanBeReEncryptedAndPreservesOpaqueStreams() {
            const string sourcePassword = "source-pass";
            const string targetPassword = "target-pass";
            byte[] plainBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Original secret", 100000,
                    100000, 2200000, 600000);
                plainBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation plain = LegacyPptPresentation.Load(plainBytes);
            byte[] opaqueBytes = "opaque-vendor-stream"u8.ToArray();
            byte[] withOpaque = plain.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]> {
                    ["VendorData/opaque"] = opaqueBytes
                });
            byte[] encryptedSource = LegacyPptRc4CryptoApi.EncryptPackage(
                withOpaque, sourcePassword);

            byte[] encryptedEdited;
            using (var input = new MemoryStream(encryptedSource))
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.LoadEncrypted(input,
                       sourcePassword)) {
                Assert.Single(presentation.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Original secret").Text =
                    "Edited secret";
                encryptedEdited = presentation.ToEncryptedBytes(
                    targetPassword, PowerPointFileFormat.Ppt);
            }

            Assert.Throws<CryptographicException>(() =>
                LegacyPptPresentation.Load(encryptedEdited,
                    new LegacyPptImportOptions { Password = sourcePassword }));
            LegacyPptPresentation edited = LegacyPptPresentation.Load(
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
                using (PowerPointPresentation source =
                       PowerPointPresentation.Create()) {
                    source.AddSlide().AddTextBox("Encrypted variant");
                    source.SaveEncrypted(path, "variant-pass");
                }

                using PowerPointPresentation loaded =
                    PowerPointPresentation.LoadEncrypted(path,
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
                using (PowerPointPresentation source =
                       PowerPointPresentation.Create()) {
                    source.AddSlide().AddTextBox("Async encrypted binary");
                    source.SaveEncrypted(path, password);
                }

                using PowerPointPresentation fromPath =
                    await PowerPointPresentation.LoadEncryptedAsync(path,
                        password);
                Assert.Equal(PowerPointFileFormat.Pps,
                    fromPath.SourceFormat);

                byte[] bytes = File.ReadAllBytes(path);
                using var encryptedStream = new MemoryStream(bytes);
                using PowerPointPresentation fromStream =
                    await PowerPointPresentation.LoadEncryptedAsync(
                        encryptedStream, password);
                Assert.Equal(PowerPointFileFormat.Ppt,
                    fromStream.SourceFormat);
                Assert.Contains(fromStream.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Async encrypted binary");

                using var normalStream = new MemoryStream(bytes);
                using PowerPointPresentation normalFlow =
                    await PowerPointPresentation.LoadAsync(normalStream,
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
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Cancellation boundary");
                encrypted = source.ToEncryptedBytes(password, format);
            }

            using var cancellation = new CancellationTokenSource();
            using var input = new CancelAtEndAsyncReadStream(encrypted,
                cancellation.Cancel);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                PowerPointPresentation.LoadEncryptedAsync(input, password,
                    cancellationToken: cancellation.Token));
        }

        [Fact]
        public void LegacyEncryption_StreamSaveRetainsImportedBinaryFormat() {
            const string sourcePassword = "stream-source";
            const string targetPassword = "stream-target";
            byte[] encrypted = CreateEncryptedLegacyPresentation(
                sourcePassword);

            using var input = new MemoryStream(encrypted);
            using PowerPointPresentation presentation =
                PowerPointPresentation.LoadEncrypted(input, sourcePassword);
            Assert.Equal(PowerPointFileFormat.Ppt,
                presentation.SourceFormat);
            presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "Original secret").Text =
                    "Re-encrypted stream";

            using var output = new MemoryStream();
            presentation.SaveEncrypted(output, targetPassword);
            using var reopenedInput = new MemoryStream(output.ToArray());
            using PowerPointPresentation reopened =
                PowerPointPresentation.LoadEncrypted(reopenedInput,
                    targetPassword);
            Assert.Contains(reopened.Slides[0].TextBoxes,
                textBox => textBox.Text == "Re-encrypted stream");
        }

        [Theory]
        [InlineData(32)]
        [InlineData(41)]
        [InlineData(136)]
        public void LegacyEncryption_RejectsInvalidRc4KeySizes(int keySizeBits) {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            source.AddSlide().AddTextBox("Invalid key size");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                source.ToEncryptedBytes("key-pass", PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LegacyPptEncryptionKeySizeBits = keySizeBits
                    }));
        }

        private static byte[] CreateEncryptedLegacyPresentation(
            string password) {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            source.AddSlide().AddTextBox("Original secret", 100000,
                100000, 2200000, 600000);
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
