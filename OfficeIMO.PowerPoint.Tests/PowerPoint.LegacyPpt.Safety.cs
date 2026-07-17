using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Compression;
using System.Threading.Tasks;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
        public void PackageReader_ChargesPersistIdsToRecordCountBudget() {
            const string Password = "persist-budget";
            byte[] bytes;
            byte[] encrypted;
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create()) {
                for (int index = 0; index < 32; index++) {
                    presentation.AddSlide().AddTextBox(
                        "Persist budget " + index);
                }
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
                encrypted = presentation.ToEncryptedBytes(Password,
                    PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation source = LegacyPptPresentation.Load(bytes);
            int persistIdCount = source.Package.UserEdits.Sum(edit =>
                edit.PersistObjectOffsets.Count);
            Assert.True(persistIdCount > 16);

            InvalidDataException plain = Assert.Throws<InvalidDataException>(
                () => LegacyPptPackage.Read(bytes,
                    new LegacyPptImportOptions { MaxRecordCount = 16 }));

            Assert.Contains("record count", plain.Message,
                StringComparison.OrdinalIgnoreCase);

            InvalidDataException encryptedException = Assert.Throws<
                InvalidDataException>(() => LegacyPptPresentation.Load(
                encrypted, new LegacyPptImportOptions {
                    Password = Password,
                    MaxRecordCount = 16
                }));
            Assert.Contains("record count", encryptedException.Message,
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
        public async Task PresentationFacade_EnforcesBinaryInputBudget() {
            byte[] bytes = CreatePresentationBytes();
            var loadOptions = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    MaxInputBytes = bytes.Length - 1
                }
            };

            using var loadStream = new MemoryStream(bytes, writable: false);
            using var encryptedStream = new MemoryStream(bytes,
                writable: false);
            using var loadAsyncStream = new MemoryStream(bytes,
                writable: false);
            using var encryptedAsyncStream = new MemoryStream(bytes,
                writable: false);

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

            var paddedBytes = new byte[256 * 1024];
            Buffer.BlockCopy(bytes, 0, paddedBytes, 0, bytes.Length);
            using var nonSeekable = new CountingNonSeekableReadStream(
                paddedBytes);
            Assert.Throws<InvalidDataException>(() =>
                PowerPointPresentation.Load(nonSeekable, loadOptions));
            Assert.Equal(paddedBytes.Length, nonSeekable.BytesRead);
        }

        [Theory]
        [InlineData(".ppt")]
        [InlineData(".pot")]
        [InlineData(".pps")]
        public async Task PresentationFacade_EnforcesLegacyExtensionLimitOnMalformedInput(
            string extension) {
            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid().ToString("N") + extension);
            try {
                File.WriteAllBytes(path, new byte[65]);
                var options = new PowerPointLoadOptions {
                    LegacyPptImportOptions = new LegacyPptImportOptions {
                        MaxInputBytes = 64
                    }
                };

                InvalidDataException syncException = Assert.Throws<
                    InvalidDataException>(() => PowerPointPresentation.Load(
                    path, options));
                InvalidDataException asyncException = await Assert
                    .ThrowsAsync<InvalidDataException>(() =>
                        PowerPointPresentation.LoadAsync(path, options));

                Assert.Contains("maximum size", syncException.Message,
                    StringComparison.OrdinalIgnoreCase);
                Assert.Contains("maximum size", asyncException.Message,
                    StringComparison.OrdinalIgnoreCase);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void PresentationFacade_EnforcesCompoundTemporaryStorageBudget() {
            byte[] binary = CreatePresentationBytes();
            var padded = new byte[256 * 1024];
            Buffer.BlockCopy(binary, 0, padded, 0, binary.Length);
            var options = new PowerPointLoadOptions {
                PackageSecurity = new OfficePackageSecurityOptions {
                    MaxPackageBytes = 4096
                }
            };
            using var input = new CountingNonSeekableReadStream(padded);

            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(
                    input, options));

            Assert.Equal(OfficePackageSecurityRule.PackageSize,
                exception.Rule);
            Assert.Contains("4096", exception.Message,
                StringComparison.Ordinal);
            Assert.Equal(4097, input.BytesRead);
        }

        [Fact]
        public void PresentationFacade_RejectsOversizedSeekableCompoundBeforeDetection() {
            byte[] binary = CreatePresentationBytes();
            var options = new PowerPointLoadOptions {
                PackageSecurity = new OfficePackageSecurityOptions {
                    MaxPackageBytes = binary.Length - 1L
                }
            };
            using var input = new CountingSeekableReadStream(binary);

            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(
                    input, options));

            Assert.Equal(OfficePackageSecurityRule.PackageSize,
                exception.Rule);
            Assert.Contains((binary.Length - 1L).ToString(),
                exception.Message, StringComparison.Ordinal);
            Assert.Equal(0, input.BytesRead);
        }

#if NET8_0_OR_GREATER
        [Fact]
        public void PresentationFacade_TemporaryStorageIsOwnerOnlyOnUnix() {
            if (OperatingSystem.IsWindows()) return;
            using FileStream temporary = PowerPointPresentation
                .CreateTemporaryInputStream(useAsync: false);
            const UnixFileMode accessBits = UnixFileMode.UserRead
                | UnixFileMode.UserWrite
                | UnixFileMode.UserExecute
                | UnixFileMode.GroupRead
                | UnixFileMode.GroupWrite
                | UnixFileMode.GroupExecute
                | UnixFileMode.OtherRead
                | UnixFileMode.OtherWrite
                | UnixFileMode.OtherExecute;

            Assert.False(File.Exists(temporary.Name));
            string descriptorPath = "/dev/fd/"
                + temporary.SafeFileHandle.DangerousGetHandle()
                    .ToInt64();
            Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite,
                File.GetUnixFileMode(descriptorPath) & accessBits);
            temporary.WriteByte(0x5A);
            temporary.Position = 0;
            Assert.Equal(0x5A, temporary.ReadByte());
        }
#endif

        [Fact]
        public async Task PresentationFacade_EnforcesExplicitPackageInputBudgetForOpenXml() {
            byte[] packageBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Package input budget");
                packageBytes = source.ToBytes();
            }
            long maxBytes = packageBytes.Length - 1L;
            var options = new PowerPointLoadOptions {
                PackageSecurity = new OfficePackageSecurityOptions {
                    MaxPackageBytes = maxBytes
                }
            };

            using var syncInput = new MemoryStream(packageBytes,
                writable: false);
            OfficePackageSecurityException syncException = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(syncInput, options));
            Assert.Equal(OfficePackageSecurityRule.PackageSize,
                syncException.Rule);

            using var asyncInput = new MemoryStream(packageBytes,
                writable: false);
            OfficePackageSecurityException asyncException = await Assert
                .ThrowsAsync<OfficePackageSecurityException>(() =>
                PowerPointPresentation.LoadAsync(asyncInput, options));
            Assert.Equal(OfficePackageSecurityRule.PackageSize,
                asyncException.Rule);

            using var nonSeekable = new CountingNonSeekableReadStream(
                packageBytes);
            OfficePackageSecurityException nonSeekableException = Assert
                .Throws<OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(nonSeekable, options));
            Assert.Equal(OfficePackageSecurityRule.PackageSize,
                nonSeekableException.Rule);
            Assert.Equal(maxBytes + 1, nonSeekable.BytesRead);

            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");
            try {
                File.WriteAllBytes(path, packageBytes);
                OfficePackageSecurityException pathException = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(path, options));
                Assert.Equal(OfficePackageSecurityRule.PackageSize,
                    pathException.Rule);
                OfficePackageSecurityException asyncPathException = await
                    Assert.ThrowsAsync<OfficePackageSecurityException>(() =>
                        PowerPointPresentation.LoadAsync(path, options));
                Assert.Equal(OfficePackageSecurityRule.PackageSize,
                    asyncPathException.Rule);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task EncryptedPresentationFacade_EnforcesTypedPackageInputBudget() {
            const string password = "typed-package-limit";
            byte[] encrypted;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Encrypted package budget");
                encrypted = source.ToEncryptedBytes(password);
            }
            var options = new PowerPointLoadOptions {
                PackageSecurity = new OfficePackageSecurityOptions {
                    MaxPackageBytes = encrypted.Length - 1L
                }
            };

            using (var syncInput = new MemoryStream(encrypted,
                       writable: false)) {
                OfficePackageSecurityException exception = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                    PowerPointPresentation.LoadEncrypted(syncInput,
                        password, options));
                Assert.Equal(OfficePackageSecurityRule.PackageSize,
                    exception.Rule);
            }
            using (var asyncInput = new MemoryStream(encrypted,
                       writable: false)) {
                OfficePackageSecurityException exception = await Assert
                    .ThrowsAsync<OfficePackageSecurityException>(() =>
                        PowerPointPresentation.LoadEncryptedAsync(asyncInput,
                            password, options));
                Assert.Equal(OfficePackageSecurityRule.PackageSize,
                    exception.Rule);
            }
            using (var nonSeekable = new CountingNonSeekableReadStream(
                       encrypted)) {
                OfficePackageSecurityException exception = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                    PowerPointPresentation.LoadEncrypted(nonSeekable,
                        password, options));
                Assert.Equal(OfficePackageSecurityRule.PackageSize,
                    exception.Rule);
                Assert.Equal(encrypted.Length, nonSeekable.BytesRead);
            }

            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");
            try {
                File.WriteAllBytes(path, encrypted);
                OfficePackageSecurityException pathException = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                    PowerPointPresentation.LoadEncrypted(path, password,
                        options));
                Assert.Equal(OfficePackageSecurityRule.PackageSize,
                    pathException.Rule);
                OfficePackageSecurityException asyncPathException = await
                    Assert.ThrowsAsync<OfficePackageSecurityException>(() =>
                        PowerPointPresentation.LoadEncryptedAsync(path,
                            password, options));
                Assert.Equal(OfficePackageSecurityRule.PackageSize,
                    asyncPathException.Rule);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task PresentationFacade_DoesNotApplyBinaryBudgetToOpenXml() {
            const string password = "openxml-budget";
            byte[] packageBytes;
            byte[] encryptedBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Open XML remains unbounded");
                packageBytes = source.ToBytes();
                encryptedBytes = source.ToEncryptedBytes(password);
            }
            var packageOptions = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    MaxInputBytes = packageBytes.Length - 1
                }
            };
            var encryptedOptions = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    MaxInputBytes = encryptedBytes.Length - 1
                }
            };

            using var packageInput = new MemoryStream(packageBytes,
                writable: false);
            using PowerPointPresentation package =
                PowerPointPresentation.Load(packageInput, packageOptions);
            Assert.Single(package.Slides);

            using var packageAsyncInput = new MemoryStream(packageBytes,
                writable: false);
            using PowerPointPresentation packageAsync =
                await PowerPointPresentation.LoadAsync(packageAsyncInput,
                    packageOptions);
            Assert.Single(packageAsync.Slides);

            using var encryptedInput = new MemoryStream(encryptedBytes,
                writable: false);
            using PowerPointPresentation encrypted =
                PowerPointPresentation.LoadEncrypted(encryptedInput,
                    password, encryptedOptions);
            Assert.Single(encrypted.Slides);

            using var encryptedNonSeekable =
                new CountingNonSeekableReadStream(encryptedBytes);
            using PowerPointPresentation encryptedFromNonSeekable =
                PowerPointPresentation.LoadEncrypted(encryptedNonSeekable,
                    password, encryptedOptions);
            Assert.Single(encryptedFromNonSeekable.Slides);

            using var encryptedAsyncInput = new MemoryStream(encryptedBytes,
                writable: false);
            using PowerPointPresentation encryptedAsync =
                await PowerPointPresentation.LoadEncryptedAsync(
                    encryptedAsyncInput, password, encryptedOptions);
            Assert.Single(encryptedAsync.Slides);
        }

        [Fact]
        public async Task PresentationFacade_ClassifiesCompleteLargeEncryptedOpenXmlFromNonSeekableStream() {
            const string password = "large-openxml-budget";
            var randomBytes = new byte[128 * 1024];
            new Random(42).NextBytes(randomBytes);
            string largeText = Convert.ToBase64String(randomBytes);
            byte[] encryptedBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox(largeText);
                encryptedBytes = source.ToEncryptedBytes(password);
            }

            Assert.True(encryptedBytes.Length > 81920);
            var options = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    MaxInputBytes = 4096
                }
            };

            using var syncInput = new CountingNonSeekableReadStream(
                encryptedBytes);
            using PowerPointPresentation sync =
                PowerPointPresentation.LoadEncrypted(syncInput, password,
                    options);
            Assert.Equal(largeText.Length,
                Assert.Single(sync.Slides[0].TextBoxes).Text.Length);
            Assert.Equal(encryptedBytes.Length, syncInput.BytesRead);

            using var asyncInput = new CountingNonSeekableReadStream(
                encryptedBytes);
            using PowerPointPresentation asyncPresentation =
                await PowerPointPresentation.LoadEncryptedAsync(asyncInput,
                    password, options);
            Assert.Equal(largeText.Length,
                Assert.Single(asyncPresentation.Slides[0].TextBoxes)
                    .Text.Length);
            Assert.Equal(encryptedBytes.Length, asyncInput.BytesRead);
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
        public void CompoundStorageValidation_BoundsOleAndVbaLogicalExpansion() {
            var options = new LegacyPptImportOptions();
            byte[] oleStorage = CreateOleTestStorage("Bounded import OLE");
            foreach (string streamName in new[] {
                         "\u0001Ole10Native", "CONTENTS"
                     }) {
                int entry = FindCompoundDirectoryEntry(oleStorage,
                    streamName);
                WriteCompoundUInt64(oleStorage, entry + 120,
                    checked((ulong)oleStorage.Length));
            }

            Assert.False(LegacyPptCompoundStorageValidator.TryRead(
                oleStorage, options, out _, out string? oleReason));
            Assert.Contains("Compound stream bytes exceed", oleReason,
                StringComparison.OrdinalIgnoreCase);

            byte[] vbaStorage = CreateVbaTestProject("BoundedModule",
                "Sub Main(): End Sub");
            foreach (string streamName in new[] {
                         "dir", "_VBA_PROJECT"
                     }) {
                int entry = FindCompoundDirectoryEntry(vbaStorage,
                    streamName);
                WriteCompoundUInt64(vbaStorage, entry + 120,
                    checked((ulong)vbaStorage.Length));
            }

            Assert.False(LegacyPptVbaProjectCodec.IsValidProject(
                vbaStorage, options, out string? vbaReason));
            Assert.Contains("Compound stream bytes exceed", vbaReason,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task PresentationFacade_EnforcesPackageSecurityPolicies() {
            byte[] packageBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Security policy");
                packageBytes = source.ToBytes();
            }
            using (var editable = new MemoryStream()) {
                editable.Write(packageBytes, 0, packageBytes.Length);
                editable.Position = 0;
                using (PresentationDocument document =
                       PresentationDocument.Open(editable, true)) {
                    document.PresentationPart!.AddExternalRelationship(
                        "urn:officeimo:test", new Uri(
                            "https://example.test/presentation"),
                        "rSecurityExternal");
                }
                packageBytes = editable.ToArray();
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using (var input = new MemoryStream(packageBytes,
                       writable: false)) {
                OfficePackageSecurityException exception = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                        PowerPointPresentation.Load(input, loadOptions));
                Assert.Equal(OfficePackageSecurityRule
                    .ExternalRelationships, exception.Rule);
            }
            using (var input = new MemoryStream(packageBytes,
                       writable: false)) {
                OfficePackageSecurityException exception = await Assert
                    .ThrowsAsync<OfficePackageSecurityException>(() =>
                        PowerPointPresentation.LoadAsync(input,
                            loadOptions));
                Assert.Equal(OfficePackageSecurityRule
                    .ExternalRelationships, exception.Rule);
            }

            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");
            try {
                File.WriteAllBytes(path, packageBytes);
                OfficePackageSecurityException exception = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                        PowerPointPresentation.Load(path, loadOptions));
                Assert.Equal(OfficePackageSecurityRule
                    .ExternalRelationships, exception.Rule);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void PresentationFacade_EnforcesPackageSecurityOnLegacyVba() {
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Legacy VBA security policy");
                SetVbaProject(source, CreateVbaTestProject(
                    "SecurityModule", "Sub Main(): End Sub"));
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.Macros, exception.Rule);
        }

        [Fact]
        public void LegacyVbaConversion_EnforcesPackageSecurityBeforeOpeningPackage() {
            byte[] packageBytes;
            using (var package = new MemoryStream()) {
                using (var archive = new ZipArchive(package,
                           ZipArchiveMode.Create, leaveOpen: true)) {
                    ZipArchiveEntry vba = archive.CreateEntry(
                        "ppt/vbaProject.bin");
                    using Stream payload = vba.Open();
                    payload.WriteByte(1);
                }
                packageBytes = package.ToArray();
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation
                        .ConvertProjectedVbaPackageToMacroEnabled(
                            packageBytes, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.Macros, exception.Rule);
        }

        [Fact]
        public void PresentationFacade_EnforcesPackageSecurityOnOriginalLegacyContainer() {
            byte[] binary = CreatePresentationBytes();
            LegacyPptPresentation source = LegacyPptPresentation.Load(binary);
            byte[] withOpaqueObjectPool = source.Package
                .RewriteCompoundStreams(new Dictionary<string, byte[]> {
                    ["ObjectPool/Preserved/Contents"] =
                        new byte[] { 1, 2, 3, 4 }
                });
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using var input = new MemoryStream(withOpaqueObjectPool,
                writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.EmbeddedPayloads,
                exception.Rule);
        }

        [Theory]
        [InlineData(false, OfficePackageSecurityRule.EmbeddedPayloads)]
        [InlineData(true, OfficePackageSecurityRule.ActiveX)]
        public void PresentationFacade_EnforcesLegacyActiveContentPolicies(
            bool activeX, OfficePackageSecurityRule expectedRule) {
            byte[] storage = CreateOleTestStorage(activeX
                ? "ActiveX policy"
                : "OLE policy");
            byte[] binary;
            if (activeX) {
                binary = CreateExternalObjectFixture(storage,
                    ExternalObjectFixtureKind.ActiveX, compressed: false);
            } else {
                using PowerPointPresentation source =
                    PowerPointPresentation.Create();
                PowerPointSlide slide = source.AddSlide();
                using var payload = new MemoryStream(storage,
                    writable: false);
                slide.AddOleObject(payload, "Package");
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(expectedRule, exception.Rule);
        }

        [Fact]
        public void PresentationFacade_RejectsPreserveOnlyLegacyExternalContent() {
            byte[] storage = CreateOleTestStorage(
                "Preserve-only linked OLE policy");
            byte[] binary = CreateExternalObjectFixture(storage,
                ExternalObjectFixtureKind.LinkedOle, compressed: false,
                linkedUpdateMode: uint.MaxValue);
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                binary);
            Assert.Empty(neutral.LinkedOleObjects);
            Assert.Contains(neutral.Diagnostics, diagnostic =>
                diagnostic.Code.StartsWith("PPT-OLE-LINK-",
                    StringComparison.Ordinal));
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = security
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.ExternalRelationships,
                exception.Rule);
        }

        [Fact]
        public void LegacyExternalPolicy_RecognizesLocationOnlyHyperlinks() {
            var hyperlink = new LegacyPptHyperlink(1, friendlyName: null,
                target: null, location: "https://example.test/location");

            Assert.True(PowerPointPresentation.IsExternalLegacyHyperlink(
                hyperlink));
        }

        [Fact]
        public void PresentationFacade_RejectsLegacyRunProgramActions() {
            var programUri = new Uri("file:///Applications/Calculator.app");
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(programUri, true);
                P.NonVisualDrawingProperties properties =
                    ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!;
                properties.Append(new A.HyperlinkOnClick {
                    Id = relationship.Id,
                    Action = "ppaction://program"
                });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                binary);
            Assert.Contains(neutral.Slides[0].Shapes.SelectMany(shape =>
                    shape.Interactions), interaction =>
                interaction.Action ==
                    LegacyPptInteractionAction.RunProgram);
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = security
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.ExternalRelationships,
                exception.Rule);
        }

        [Fact]
        public void EncryptedLegacyLoad_ValidatesOuterSourceBeforePasswordProcessing() {
            const string password = "source-policy-pass";
            byte[] encrypted;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Encrypted source policy");
                encrypted = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt);
            }
            OfficePackageSecurityReport report =
                OfficePackageSecurityInspector.Inspect(encrypted);
            Assert.Equal(OfficePackageContainerKind.CompoundBinary,
                report.ContainerKind);
            Assert.True(report.PartCount > 1);
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.MaxPartCount = report.PartCount - 1;
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = security
            };

            using var input = new MemoryStream(encrypted, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.LoadEncrypted(input,
                        "wrong-password", loadOptions));

            Assert.Equal(OfficePackageSecurityRule.PartCount,
                exception.Rule);
        }

        [Fact]
        public void PackageReader_ChargesDecodedPicturesToSharedBudget() {
            byte[] firstImage = OfficePngWriter.Encode(new OfficeRasterImage(
                4, 4, OfficeColor.FromRgb(10, 20, 30)));
            byte[] secondImage = OfficePngWriter.Encode(new OfficeRasterImage(
                4, 4, OfficeColor.FromRgb(40, 50, 60)));
            byte[] bytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide();
                using (var first = new MemoryStream(firstImage,
                           writable: false)) {
                    slide.AddPicture(first,
                        OfficeIMO.PowerPoint.ImagePartType.Png);
                }
                using (var second = new MemoryStream(secondImage,
                           writable: false)) {
                    slide.AddPicture(second,
                        OfficeIMO.PowerPoint.ImagePartType.Png,
                        left: 1000000);
                }
                bytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation unrestricted =
                LegacyPptPresentation.Load(bytes);
            Assert.Equal(2, unrestricted.BlipStoreEntries.Count);
            long decodedBytes = unrestricted.BlipStoreEntries.Sum(entry =>
                (long)entry.ImageByteCount);

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptPresentation.Load(bytes,
                new LegacyPptImportOptions {
                    MaxDecodedStorageBytes = decodedBytes - 1
                }));

            Assert.Contains("aggregate decoded embedded-storage",
                exception.Message, StringComparison.OrdinalIgnoreCase);
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
        public void EncryptedSummaryStorage_ReservesSharedBudgetBeforeDecode() {
            byte[] storage = CreateOleTestStorage(
                "Encrypted summary storage");
            var replacements = new Dictionary<string, byte[]>(
                StringComparer.OrdinalIgnoreCase);
            var options = new LegacyPptImportOptions {
                MaxDecodedStorageBytes = 1
            };
            var budget = new LegacyPptDecodedStorageBudget(
                options.MaxDecodedStorageBytes);

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptEncryptedSummary
                    .AddStorageReplacements(replacements, "Nested", storage,
                        options, budget));

            Assert.Contains("aggregate decoded embedded-storage",
                exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(replacements);
            Assert.Equal(0, budget.DecodedBytes);
            Assert.Throws<InvalidDataException>(() =>
                budget.ThrowIfExceeded());
        }

        [Fact]
        public void PackageReader_BoundsAliasedTopLevelCompoundStreams() {
            byte[] expanded = CreatePresentationBytes();
            int documentEntry = FindCompoundDirectoryEntry(expanded,
                "PowerPoint Document");
            uint documentStart = ReadCompoundUInt32(expanded,
                documentEntry + 116);
            ulong documentSize = ReadCompoundUInt64(expanded,
                documentEntry + 120);
            Assert.True(documentSize > 4096);
            foreach (string alias in new[] {
                         "Current User", "\u0005SummaryInformation",
                         "\u0005DocumentSummaryInformation"
                     }) {
                int aliasEntry = FindCompoundDirectoryEntry(expanded,
                    alias);
                WriteUInt32(expanded, aliasEntry + 116, documentStart);
                WriteCompoundUInt64(expanded, aliasEntry + 120,
                    documentSize);
            }
            Assert.True(documentSize * 4UL
                > unchecked((ulong)expanded.Length));

            var importOptions = new LegacyPptImportOptions {
                MaxInputBytes = expanded.Length
            };
            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptPresentation.Load(
                    expanded, importOptions));

            Assert.Contains("Compound stream bytes exceed",
                exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(expanded.Length.ToString(), exception.Message,
                StringComparison.Ordinal);
            using var encryptedRoute = new MemoryStream(expanded,
                writable: false);
            InvalidDataException routeException = Assert.Throws<
                InvalidDataException>(() =>
                PowerPointPresentation.LoadEncrypted(encryptedRoute,
                    "password", new PowerPointLoadOptions {
                        LegacyPptImportOptions = importOptions
                    }));
            Assert.Contains("Compound stream bytes exceed",
                routeException.Message, StringComparison.OrdinalIgnoreCase);
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

        [Fact]
        public void PackageReader_SharesPersistRecordsAliasedByOffset() {
            byte[] bytes = CreatePresentationBytes();
            LegacyPptPresentation source = LegacyPptPresentation.Load(bytes);
            byte[] document = (byte[])source.Package.DocumentStream.Clone();
            uint directoryOffset = source.Package.UserEdits[0]
                .PersistDirectoryOffset;
            LegacyPptRecord directory = LegacyPptRecordReader.ReadSingle(
                document, checked((int)directoryOffset),
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
            (uint PersistId, int ValueOffset, uint StreamOffset) first =
                entries[0];
            (uint PersistId, int ValueOffset, uint StreamOffset) second =
                entries.First(entry => entry.StreamOffset !=
                    first.StreamOffset);
            WriteUInt32(document, second.ValueOffset, first.StreamOffset);
            byte[] aliased = source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(
                    StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = document
                });

            LegacyPptPackage package = LegacyPptPackage.Read(aliased,
                new LegacyPptImportOptions());

            Assert.Same(package.PersistObjects[first.PersistId].RecordBytes,
                package.PersistObjects[second.PersistId].RecordBytes);
        }

        [Fact]
        public void PackageReader_RejectsOverlappingPersistRecordRanges() {
            byte[] bytes = CreatePresentationBytes();
            LegacyPptPresentation source = LegacyPptPresentation.Load(bytes);
            byte[] document = (byte[])source.Package.DocumentStream.Clone();
            LegacyPptRecord directory = LegacyPptRecordReader.ReadSingle(
                document, checked((int)source.Package.UserEdits[0]
                    .PersistDirectoryOffset), new LegacyPptImportOptions());
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
            (uint PersistId, int ValueOffset, uint StreamOffset) owner =
                entries.First(entry => entry.StreamOffset <= int.MaxValue
                    && ReadCompoundUInt32(document,
                        checked((int)entry.StreamOffset + 4)) >= 8);
            (uint PersistId, int ValueOffset, uint StreamOffset) alias =
                entries.First(entry => entry.PersistId != owner.PersistId);
            uint overlappingOffset = checked(owner.StreamOffset + 8U);
            int overlappingHeader = checked((int)overlappingOffset);
            WriteUInt16(document, overlappingHeader, 0);
            WriteUInt16(document, overlappingHeader + 2, 0x1000);
            WriteUInt32(document, overlappingHeader + 4, 0);
            WriteUInt32(document, alias.ValueOffset, overlappingOffset);
            byte[] overlapping = source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(
                    StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = document
                });

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptPackage.Read(
                overlapping, new LegacyPptImportOptions()));

            Assert.Contains("overlap", exception.Message,
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

        private static int FindCompoundDirectoryEntry(byte[] bytes,
            string name) {
            byte[] encoded = Encoding.Unicode.GetBytes(name + '\0');
            for (int offset = 512; offset <= bytes.Length - encoded.Length;
                 offset += 128) {
                if (bytes.AsSpan(offset, encoded.Length)
                    .SequenceEqual(encoded)) return offset;
            }
            throw new InvalidDataException(
                $"The compound directory entry '{name}' was not found.");
        }

        private static uint ReadCompoundUInt32(byte[] bytes, int offset) =>
            unchecked((uint)(bytes[offset]
                | bytes[offset + 1] << 8
                | bytes[offset + 2] << 16
                | bytes[offset + 3] << 24));

        private static ulong ReadCompoundUInt64(byte[] bytes, int offset) =>
            ReadCompoundUInt32(bytes, offset)
            | unchecked((ulong)ReadCompoundUInt32(bytes, offset + 4) << 32);

        private static void WriteCompoundUInt64(byte[] bytes, int offset,
            ulong value) {
            WriteUInt32(bytes, offset, unchecked((uint)value));
            WriteUInt32(bytes, offset + 4,
                unchecked((uint)(value >> 32)));
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

        private sealed class CountingNonSeekableReadStream : Stream {
            private readonly MemoryStream _inner;

            public CountingNonSeekableReadStream(byte[] bytes) {
                _inner = new MemoryStream(bytes, writable: false);
            }

            public int BytesRead { get; private set; }
            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override int Read(byte[] buffer, int offset, int count) {
                int read = _inner.Read(buffer, offset, count);
                BytesRead += read;
                return read;
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) =>
                throw new NotSupportedException();
            public override void SetLength(long value) =>
                throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset,
                int count) => throw new NotSupportedException();

            protected override void Dispose(bool disposing) {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }
        }

        private sealed class CountingSeekableReadStream : MemoryStream {
            internal CountingSeekableReadStream(byte[] bytes)
                : base(bytes, writable: false) { }

            internal int BytesRead { get; private set; }

            public override int Read(byte[] buffer, int offset,
                int count) {
                int read = base.Read(buffer, offset, count);
                BytesRead += read;
                return read;
            }
        }
    }
}
