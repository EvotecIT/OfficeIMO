using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class OfficeEncryptionTests {
        private const string Password = "OfficeIMO-Secret-123";

        [Fact]
        public void DecryptPackage_RejectsDeclaredPlaintextSizeBeforeAllocation() {
            byte[] plaintext = new byte[4096];
            byte[] encrypted = OfficeEncryption.EncryptPackage(plaintext,
                Password);

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => OfficeEncryption.DecryptPackage(
                encrypted, Password, CancellationToken.None,
                maximumDecryptedPackageBytes: plaintext.Length - 1L));

            Assert.Contains("4096", exception.Message,
                StringComparison.Ordinal);
            Assert.Contains("4095", exception.Message,
                StringComparison.Ordinal);
        }

        [Fact]
        public void DecryptPackage_BoundsAliasedCompoundStreams() {
            byte[] encrypted = OfficeEncryption.EncryptPackage(
                new byte[16 * 1024], Password);
            int encryptionInfoEntry = FindCompoundDirectoryEntry(encrypted,
                "EncryptionInfo");
            int encryptedPackageEntry = FindCompoundDirectoryEntry(encrypted,
                "EncryptedPackage");
            uint packageStart = ReadUInt32(encrypted,
                encryptedPackageEntry + 116);
            ulong packageSize = ReadUInt64(encrypted,
                encryptedPackageEntry + 120);
            Assert.True(packageSize > 4096);
            WriteUInt32(encrypted, encryptionInfoEntry + 116, packageStart);
            WriteUInt64(encrypted, encryptionInfoEntry + 120, packageSize);
            Assert.True(packageSize * 2UL > unchecked((ulong)encrypted.Length));

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => OfficeEncryption.DecryptPackage(
                encrypted, Password, CancellationToken.None,
                maximumDecryptedPackageBytes: 16 * 1024));

            Assert.Contains("Compound stream bytes exceed",
                exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Word_SaveEncrypted_And_LoadEncrypted_RoundTrips() {
            string path = CreateTempPath(".docx");

            using (var document = WordDocument.Create()) {
                document.AddParagraph("Encrypted Word content");
                document.SaveEncrypted(path, Password);
            }

            AssertEncryptedContainer(path);
            Assert.ThrowsAny<Exception>(() => WordprocessingDocument.Open(path, false).Dispose());

            using var loaded = WordDocument.LoadEncrypted(path, Password);
            Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text == "Encrypted Word content");
        }

        [Fact]
        public void Word_SaveEncryptedStream_And_LoadEncryptedStream_RoundTrips() {
            using var encrypted = new MemoryStream();

            using (var document = WordDocument.Create()) {
                document.AddParagraph("Encrypted Word stream content");
                document.SaveEncrypted(encrypted, Password);
            }

            AssertEncryptedContainer(encrypted);

            encrypted.Position = 0;
            using var loaded = WordDocument.LoadEncrypted(encrypted, Password);
            Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text == "Encrypted Word stream content");
        }

        [Fact]
        public void Word_LoadEncrypted_DoesNotAttachEncryptedPathOrAllowAutoSave() {
            string path = CreateTempPath(".docx");

            using (var document = WordDocument.Create()) {
                document.AddParagraph("Encrypted Word content");
                document.SaveEncrypted(path, Password);
            }

            using (var loaded = WordDocument.LoadEncrypted(path, Password)) {
                Assert.True(string.IsNullOrEmpty(loaded.FilePath));
                loaded.AddParagraph("Explicit encrypted edit");
                Assert.Throws<InvalidOperationException>(() => loaded.Save());
            }

            Assert.Throws<NotSupportedException>(() => WordDocument.LoadEncrypted(path, Password, new WordLoadOptions {
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            }));

            using var explicitLoad = WordDocument.LoadEncrypted(path, Password, new WordLoadOptions {
                OpenSettings = new OpenSettings { AutoSave = true }
            });
            Assert.Equal(OfficeIMO.Drawing.DocumentPersistenceMode.Explicit, explicitLoad.PersistenceMode);
        }

        [Fact]
        public void Excel_SaveEncrypted_And_LoadEncrypted_RoundTrips() {
            string path = CreateTempPath(".xlsx");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Encrypted");
                sheet.CellValue(1, 1, "Encrypted Excel content");
                document.SaveEncrypted(path, Password);
            }

            AssertEncryptedContainer(path);
            Assert.ThrowsAny<Exception>(() => SpreadsheetDocument.Open(path, false).Dispose());

            using var loaded = ExcelDocument.LoadEncrypted(path, Password);
            Assert.Equal("Encrypted", loaded.Sheets[0].Name);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out var value));
            Assert.Equal("Encrypted Excel content", value);
        }

        [Fact]
        public void Excel_SaveEncryptedStream_And_LoadEncryptedStream_RoundTrips() {
            using var encrypted = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("EncryptedStream");
                sheet.CellValue(1, 1, "Encrypted Excel stream content");
                document.SaveEncrypted(encrypted, Password);
            }

            AssertEncryptedContainer(encrypted);

            encrypted.Position = 0;
            using var loaded = ExcelDocument.LoadEncrypted(encrypted, Password);
            Assert.Equal("EncryptedStream", loaded.Sheets[0].Name);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out var value));
            Assert.Equal("Encrypted Excel stream content", value);
        }

        [Fact]
        public void Excel_LoadEncrypted_DoesNotAttachEncryptedPathOrAllowAutoSave() {
            string path = CreateTempPath(".xlsx");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Encrypted");
                sheet.CellValue(1, 1, "Encrypted Excel content");
                document.SaveEncrypted(path, Password);
            }

            using (var loaded = ExcelDocument.LoadEncrypted(path, Password)) {
                Assert.True(string.IsNullOrEmpty(loaded.FilePath));
            }

            Assert.Throws<NotSupportedException>(() => ExcelDocument.LoadEncrypted(path, Password, new ExcelLoadOptions {
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            }));

            using var explicitLoad = ExcelDocument.LoadEncrypted(path, Password, new ExcelLoadOptions {
                OpenSettings = new OpenSettings { AutoSave = true }
            });
            Assert.Equal(OfficeIMO.Drawing.DocumentPersistenceMode.Explicit, explicitLoad.PersistenceMode);
        }

        [Fact]
        public void PowerPoint_SaveEncrypted_And_OpenEncrypted_RoundTrips() {
            string path = CreateTempPath(".pptx");

            using (var presentation = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions())) {
                var slide = presentation.AddSlide();
                slide.AddTextBox("Encrypted PowerPoint content", 1, 1, 4, 1);
                presentation.SaveEncrypted(path, Password);
            }

            AssertEncryptedContainer(path);
            Assert.ThrowsAny<Exception>(() => PresentationDocument.Open(path, false).Dispose());

            using var loaded = PowerPointPresentation.LoadEncrypted(path, Password);
            Assert.Single(loaded.Slides);
        }

        [Fact]
        public void PowerPoint_SaveEncryptedStream_And_OpenEncryptedStream_RoundTrips() {
            using var encrypted = new MemoryStream();

            using (var presentation = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions())) {
                var slide = presentation.AddSlide();
                slide.AddTextBox("Encrypted PowerPoint stream content", 1, 1, 4, 1);
                presentation.SaveEncrypted(encrypted, Password);
            }

            AssertEncryptedContainer(encrypted);

            encrypted.Position = 0;
            using var loaded = PowerPointPresentation.LoadEncrypted(encrypted, Password);
            Assert.Single(loaded.Slides);
        }

        [Fact]
        public void Excel_LoadEncrypted_WithWrongPassword_ThrowsCryptographicException() {
            string path = CreateTempPath(".xlsx");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AddWorksheet("Encrypted");
                document.SaveEncrypted(path, Password);
            }

            Assert.Throws<CryptographicException>(() => ExcelDocument.LoadEncrypted(path, "wrong-password"));
        }

        [Fact]
        public void Excel_LoadEncrypted_WithTamperedPayload_ThrowsCryptographicException() {
            string path = CreateTempPath(".xlsx");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Encrypted");
                sheet.CellValue(1, 1, "Tamper target");
                document.SaveEncrypted(path, Password);
            }

            byte[] bytes = File.ReadAllBytes(path);
            bytes[512 + 100] ^= 0xff;
            File.WriteAllBytes(path, bytes);

            Assert.Throws<CryptographicException>(() => ExcelDocument.LoadEncrypted(path, Password));
        }

        private static string CreateTempPath(string extension) {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N") + extension);
            return path;
        }

        private static void AssertEncryptedContainer(string path) {
            AssertEncryptedContainer(File.ReadAllBytes(path));
        }

        private static void AssertEncryptedContainer(MemoryStream stream) {
            AssertEncryptedContainer(stream.ToArray());
        }

        private static void AssertEncryptedContainer(byte[] bytes) {
            Assert.True(bytes.Length > 512);
            Assert.Equal(0xd0, bytes[0]);
            Assert.Equal(0xcf, bytes[1]);
            Assert.Equal(0x11, bytes[2]);
            Assert.Equal(0xe0, bytes[3]);
        }

        private static int FindCompoundDirectoryEntry(byte[] bytes,
            string name) {
            byte[] encoded = System.Text.Encoding.Unicode.GetBytes(name + '\0');
            for (int offset = 512; offset <= bytes.Length - encoded.Length;
                 offset += 128) {
                if (bytes.AsSpan(offset, encoded.Length)
                    .SequenceEqual(encoded)) return offset;
            }
            throw new InvalidDataException(
                $"The compound directory entry '{name}' was not found.");
        }

        private static uint ReadUInt32(byte[] bytes, int offset) =>
            unchecked((uint)(bytes[offset]
                | bytes[offset + 1] << 8
                | bytes[offset + 2] << 16
                | bytes[offset + 3] << 24));

        private static ulong ReadUInt64(byte[] bytes, int offset) =>
            ReadUInt32(bytes, offset)
            | unchecked((ulong)ReadUInt32(bytes, offset + 4) << 32);

        private static void WriteUInt32(byte[] bytes, int offset,
            uint value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
            bytes[offset + 2] = unchecked((byte)(value >> 16));
            bytes[offset + 3] = unchecked((byte)(value >> 24));
        }

        private static void WriteUInt64(byte[] bytes, int offset,
            ulong value) {
            WriteUInt32(bytes, offset, unchecked((uint)value));
            WriteUInt32(bytes, offset + 4,
                unchecked((uint)(value >> 32)));
        }
    }
}
