using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class OfficeEncryptionTests : IDisposable {
        private readonly string _root = Directory.CreateTempSubdirectory("officeimo-encryption-test-").FullName;
        public void Dispose() => Directory.Delete(_root, recursive: true);
        private const string Password = "OfficeIMO-Secret-123";

        [Fact]
        public void DecryptPackage_RejectsDeclaredPlaintextSizeBeforeAllocation() {
            byte[] plaintext = new byte[4096];
            byte[] encrypted = OfficeEncryption.EncryptPackage(plaintext,
                Password);

            OfficeInspectionResourceLimitException exception = Assert.Throws<
                OfficeInspectionResourceLimitException>(() => OfficeEncryption.DecryptPackage(
                encrypted, Password, CancellationToken.None,
                maximumDecryptedPackageBytes: plaintext.Length - 1L));

            Assert.Equal("encryption", exception.Owner);
            Assert.Equal("MaximumDecryptedPackageBytes", exception.LimitName);
            Assert.Equal(plaintext.Length - 1L, exception.Maximum);
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
                document.Save(path, new WordSaveOptions { Password = Password });
            }

            AssertEncryptedContainer(path);
            Assert.ThrowsAny<Exception>(() => WordprocessingDocument.Open(path, false).Dispose());

            using var loaded = WordDocument.Load(path, new WordLoadOptions { Password = Password });
            Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text == "Encrypted Word content");
        }

        [Fact]
        public void Word_SaveEncryptedStream_And_LoadEncryptedStream_RoundTrips() {
            using var encrypted = new MemoryStream();

            using (var document = WordDocument.Create()) {
                document.AddParagraph("Encrypted Word stream content");
                document.Save(encrypted, new WordSaveOptions { Password = Password });
            }

            AssertEncryptedContainer(encrypted);

            encrypted.Position = 0;
            using var loaded = WordDocument.Load(encrypted, new WordLoadOptions { Password = Password });
            Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text == "Encrypted Word stream content");
        }

        [Fact]
        public void Word_PasswordLoadRetainsTheAssociatedPathAndEncryptsSubsequentSaves() {
            string file = CreateTempPath(".docx");
            using (var document = WordDocument.Create()) { document.AddParagraph("Original"); document.Save(file, new WordSaveOptions { Password = Password }); }
            using (var document = WordDocument.Load(file, new WordLoadOptions { Password = Password })) {
                Assert.Equal(file, document.FilePath);
                document.AddParagraph("Explicit encrypted edit");
                document.Save();
            }
            using (var document = WordDocument.Load(file, new WordLoadOptions { Password = Password, PersistenceMode = DocumentPersistenceMode.SaveOnDispose }))
                document.AddParagraph("Automatic encrypted edit");
            AssertEncryptedContainer(file);
            using var reopened = WordDocument.Load(file, new WordLoadOptions { Password = Password });
            Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Explicit encrypted edit");
            Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Automatic encrypted edit");
        }

        [Fact]
        public void Excel_SaveEncrypted_And_LoadEncrypted_RoundTrips() {
            string path = CreateTempPath(".xlsx");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Encrypted");
                sheet.CellValue(1, 1, "Encrypted Excel content");
                document.Save(path, new ExcelSaveOptions { Password = Password });
            }

            AssertEncryptedContainer(path);
            Assert.ThrowsAny<Exception>(() => SpreadsheetDocument.Open(path, false).Dispose());

            using var loaded = ExcelDocument.Load(path, new ExcelLoadOptions { Password = Password });
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
                document.Save(encrypted, new ExcelSaveOptions { Password = Password });
            }

            AssertEncryptedContainer(encrypted);

            encrypted.Position = 0;
            using var loaded = ExcelDocument.Load(encrypted, new ExcelLoadOptions { Password = Password });
            Assert.Equal("EncryptedStream", loaded.Sheets[0].Name);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out var value));
            Assert.Equal("Encrypted Excel stream content", value);
        }

        [Fact]
        public void Excel_PasswordLoadRetainsTheAssociatedPathAndEncryptsSubsequentSaves() {
            string file = CreateTempPath(".xlsx");
            using (var document = ExcelDocument.Create()) { document.AddWorksheet("Data").CellValue(1, 1, "Original"); document.Save(file, new ExcelSaveOptions { Password = Password }); }
            using (var document = ExcelDocument.Load(file, new ExcelLoadOptions { Password = Password })) {
                Assert.Equal(file, document.FilePath);
                document.Sheets[0].CellValue(2, 1, "Explicit encrypted edit");
                document.Save();
            }
            using (var document = ExcelDocument.Load(file, new ExcelLoadOptions { Password = Password, PersistenceMode = DocumentPersistenceMode.SaveOnDispose }))
                document.Sheets[0].CellValue(3, 1, "Automatic encrypted edit");
            AssertEncryptedContainer(file);
            using var reopened = ExcelDocument.Load(file, new ExcelLoadOptions { Password = Password });
            Assert.True(reopened.Sheets[0].TryGetCellText(2, 1, out string explicitText));
            Assert.True(reopened.Sheets[0].TryGetCellText(3, 1, out string automaticText));
            Assert.Equal("Explicit encrypted edit", explicitText);
            Assert.Equal("Automatic encrypted edit", automaticText);
        }

        [Fact]
        public void PowerPoint_SaveEncrypted_And_OpenEncrypted_RoundTrips() {
            string path = CreateTempPath(".pptx");

            using (var presentation = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions())) {
                var slide = presentation.AddSlide();
                slide.AddTextBox("Encrypted PowerPoint content", 1, 1, 4, 1);
                presentation.Save(path, new PowerPointSaveOptions { Password = Password });
            }

            AssertEncryptedContainer(path);
            Assert.ThrowsAny<Exception>(() => PresentationDocument.Open(path, false).Dispose());

            using var loaded = PowerPointPresentation.Load(path, new PowerPointLoadOptions { Password = Password });
            Assert.Single(loaded.Slides);
        }

        [Fact]
        public void PowerPoint_SaveEncryptedStream_And_OpenEncryptedStream_RoundTrips() {
            using var encrypted = new MemoryStream();

            using (var presentation = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions())) {
                var slide = presentation.AddSlide();
                slide.AddTextBox("Encrypted PowerPoint stream content", 1, 1, 4, 1);
                presentation.Save(encrypted, new PowerPointSaveOptions { Password = Password });
            }

            AssertEncryptedContainer(encrypted);

            encrypted.Position = 0;
            using var loaded = PowerPointPresentation.Load(encrypted, new PowerPointLoadOptions { Password = Password });
            Assert.Single(loaded.Slides);
        }

        [Fact]
        public void Excel_LoadEncrypted_WithWrongPassword_ThrowsCryptographicException() {
            string path = CreateTempPath(".xlsx");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AddWorksheet("Encrypted");
                document.Save(path, new ExcelSaveOptions { Password = Password });
            }

            Assert.Throws<CryptographicException>(() => ExcelDocument.Load(path, new ExcelLoadOptions { Password = "wrong-password" }));
        }

        [Fact]
        public void Excel_LoadEncrypted_WithTamperedPayload_ThrowsCryptographicException() {
            string path = CreateTempPath(".xlsx");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Encrypted");
                sheet.CellValue(1, 1, "Tamper target");
                document.Save(path, new ExcelSaveOptions { Password = Password });
            }

            // Locate the ciphertext by its compound stream name; its sector
            // position changes with mini-stream and directory layout choices.
            using (var encrypted = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                using RootStorage root = RootStorage.Open(encrypted, StorageModeFlags.LeaveOpen);
                using CfbStream payload = root.OpenStream("EncryptedPackage");
                payload.Position = 9;
                int value = payload.ReadByte();
                Assert.NotEqual(-1, value);
                payload.Position = 9;
                payload.WriteByte((byte)(value ^ 0xff));
                root.Flush();
            }

            Assert.Throws<CryptographicException>(() => ExcelDocument.Load(path, new ExcelLoadOptions { Password = Password }));
        }

        private string CreateTempPath(string extension) {
            string path = Path.Combine(_root, Guid.NewGuid().ToString("N") + extension);
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
