using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class AgileIntegrityTests {
        private const string IntegrityFailureMessage = "The encrypted package integrity check failed.";

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DecryptPackage_RejectsSaltLengthMismatchBeforeUsingThePassword(bool passwordSalt) {
            byte[] encrypted = OfficeEncryption.EncryptPackage(new byte[] { 1, 2, 3 }, "test-password",
                new OfficeEncryptionOptions { SpinCount = 1 });
            byte[] changed = RewriteDescriptor(encrypted, descriptor => {
                XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
                XNamespace password = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
                XElement element = passwordSalt ? descriptor.Root!.Element(ns + "keyEncryptors")!
                    .Element(ns + "keyEncryptor")!.Element(password + "encryptedKey")! : descriptor.Root!.Element(ns + "keyData")!;
                element.SetAttributeValue("saltSize", 31);
            });
            Assert.Throws<InvalidDataException>(() => OfficeEncryption.DecryptPackage(changed, "test-password"));
        }

        [Fact]
        public void DecryptPackage_DoesNotDistinguishIntegrityPaddingFromAuthenticationFailure() {
            byte[] encrypted = ReadFixture("agile-digest-sha1.xlsx");
            byte[] invalidPadding = RewriteIntegrity(encrypted, "encryptedHmacKey", bytes => { bytes[15] ^= 1; return bytes; });
            byte[] invalidTag = RewriteIntegrity(encrypted, "encryptedHmacValue", bytes => { bytes[0] ^= 1; return bytes; });
            var paddingError = Assert.Throws<CryptographicException>(() => OfficeEncryption.DecryptPackage(invalidPadding, "hunter2"));
            var tagError = Assert.Throws<CryptographicException>(() => OfficeEncryption.DecryptPackage(invalidTag, "hunter2"));
            Assert.Equal(IntegrityFailureMessage, paddingError.Message);
            Assert.Equal(tagError.Message, paddingError.Message);
        }

        [Fact]
        public void DecryptPackage_ConsistentSaltDescriptorDoesNotRestoreThePaddingOracle() {
            byte[] encrypted = OfficeEncryption.EncryptPackage(new byte[] { 1, 2, 3 }, "test-password",
                new OfficeEncryptionOptions { SpinCount = 1 });
            var messages = new HashSet<string>(StringComparer.Ordinal);
            for (int value = 0; value < 256; value++) {
                byte[] changed = RewriteDescriptor(encrypted, descriptor => {
                    XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
                    XElement keyData = descriptor.Root!.Element(ns + "keyData")!;
                    keyData.SetAttributeValue("saltSize", 31);
                    keyData.SetAttributeValue("saltValue", Convert.ToBase64String(new byte[31]));
                    XAttribute key = descriptor.Root.Element(ns + "dataIntegrity")!.Attribute("encryptedHmacKey")!;
                    byte[] ciphertext = Convert.FromBase64String(key.Value);
                    Array.Resize(ref ciphertext, 32);
                    // Every possible last plaintext byte occurs in block two, while
                    // both salt fields remain consistent with the format's range.
                    ciphertext[15] = (byte)value;
                    key.Value = Convert.ToBase64String(ciphertext);
                });
                var error = Assert.Throws<CryptographicException>(() => OfficeEncryption.DecryptPackage(changed, "test-password"));
                messages.Add(error.Message);
            }
            Assert.Equal(IntegrityFailureMessage, Assert.Single(messages));
        }

        [Theory]
        [InlineData("sha1")]
        [InlineData("sha256")]
        [InlineData("sha384")]
        [InlineData("sha512")]
        public void DecryptPackage_VerifiesIndependentDigestSizedIntegrityKeys(string hashName) {
            byte[] expected = ReadFixture("agile-aes256-sha512.plain.xlsx");
            byte[] encrypted = ReadFixture("agile-digest-" + hashName + ".xlsx");
            Assert.Equal(expected, OfficeEncryption.DecryptPackage(encrypted, "hunter2"));
        }

        [Theory]
        [InlineData("sha1")]
        [InlineData("sha256")]
        [InlineData("sha384")]
        [InlineData("sha512")]
        public void DecryptPackage_ReadsExistingSaltSizedIntegrityKeys(string hashName) {
            byte[] expected = ReadFixture("agile-aes256-sha512.plain.xlsx");
            byte[] encrypted = ReadFixture("officeimo-salt-" + hashName + ".xlsx");
            Assert.Equal(expected, OfficeEncryption.DecryptPackage(encrypted, "test-password"));
        }

        [Theory]
        [InlineData("Sha1")]
        [InlineData("Sha256")]
        [InlineData("Sha384")]
        [InlineData("Sha512")]
        public void EncryptPackage_WritesDigestSizedIntegrityKeysAndTags(string hashName) {
            var algorithm = (OfficeEncryptionHashAlgorithm)Enum.Parse(typeof(OfficeEncryptionHashAlgorithm), hashName);
            byte[] plaintext = new byte[5003];
            for (int index = 0; index < plaintext.Length; index++) plaintext[index] = (byte)index;
            byte[] encrypted = OfficeEncryption.EncryptPackage(plaintext, "test-password",
                new OfficeEncryptionOptions { HashAlgorithm = algorithm, KeyBits = 128, SpinCount = 1 });
            XDocument descriptor = ReadDescriptor(encrypted);
            XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
            XElement keyData = descriptor.Root!.Element(ns + "keyData")!;
            int digestSize = int.Parse(keyData.Attribute("hashSize")!.Value);
            XElement integrity = descriptor.Root.Element(ns + "dataIntegrity")!;
            int paddedSize = (digestSize + 15) / 16 * 16;
            Assert.Equal(paddedSize, Convert.FromBase64String(integrity.Attribute("encryptedHmacKey")!.Value).Length);
            Assert.Equal(paddedSize, Convert.FromBase64String(integrity.Attribute("encryptedHmacValue")!.Value).Length);
            Assert.Equal(plaintext, OfficeEncryption.DecryptPackage(encrypted, "test-password"));
        }

        [Theory]
        [InlineData("sha1", false)]
        [InlineData("sha1", true)]
        [InlineData("sha256", false)]
        [InlineData("sha256", true)]
        [InlineData("sha384", false)]
        [InlineData("sha384", true)]
        [InlineData("sha512", false)]
        [InlineData("sha512", true)]
        public void DecryptPackage_RejectsTamperedPayloadAndIntegrityTags(string hashName, bool tag) {
            byte[] encrypted = ReadFixture("agile-digest-" + hashName + ".xlsx");
            byte[] tampered = tag
                ? RewriteIntegrity(encrypted, "encryptedHmacValue", bytes => { bytes[0] ^= 1; return bytes; })
                : RewritePayload(encrypted);
            Assert.Throws<CryptographicException>(() => OfficeEncryption.DecryptPackage(tampered, "hunter2"));
        }

        [Fact]
        public void DecryptPackage_StillAuthenticatesExistingSaltLayout() {
            byte[] encrypted = ReadFixture("officeimo-salt-sha512.xlsx");
            Assert.Throws<CryptographicException>(() => OfficeEncryption.DecryptPackage(
                RewritePayload(encrypted), "test-password"));
            Assert.Throws<CryptographicException>(() => OfficeEncryption.DecryptPackage(
                RewriteIntegrity(encrypted, "encryptedHmacValue", bytes => { bytes[0] ^= 1; return bytes; }),
                "test-password"));
        }

        [Fact]
        public void DecryptPackage_RejectsUnrecognizedIntegrityKeyLength() {
            byte[] encrypted = ReadFixture("agile-digest-sha512.xlsx");
            byte[] tampered = RewriteIntegrity(encrypted, "encryptedHmacKey", bytes => {
                Array.Resize(ref bytes, 48);
                return bytes;
            });
            CryptographicException error = Assert.Throws<CryptographicException>(() =>
                OfficeEncryption.DecryptPackage(tampered, "hunter2"));
            Assert.Equal(IntegrityFailureMessage, error.Message);
        }

        [Fact]
        public void DecryptPackage_RejectsNonzeroIntegrityKeyPadding() {
            byte[] encrypted = ReadFixture("agile-digest-sha1.xlsx");
            byte[] tampered = RewriteIntegrity(encrypted, "encryptedHmacKey", bytes => {
                // The previous CBC block controls byte 31, in the zero padding
                // following the 20-byte SHA-1 integrity key.
                bytes[15] ^= 1;
                return bytes;
            });
            CryptographicException error = Assert.Throws<CryptographicException>(() =>
                OfficeEncryption.DecryptPackage(tampered, "hunter2"));
            Assert.Equal(IntegrityFailureMessage, error.Message);
        }

        private static byte[] ReadFixture(string name) => File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory, "Documents", "ExcelEncryptionCorpus", name));

        private static XDocument ReadDescriptor(byte[] encrypted) {
            using var source = new MemoryStream(encrypted, writable: false);
            using RootStorage root = RootStorage.Open(source, StorageModeFlags.LeaveOpen);
            using CfbStream stream = root.OpenStream("EncryptionInfo");
            using var bytes = new MemoryStream();
            stream.CopyTo(bytes);
            byte[] descriptor = bytes.ToArray();
            return XDocument.Parse(Encoding.UTF8.GetString(descriptor, 8, descriptor.Length - 8));
        }

        private static byte[] RewriteIntegrity(byte[] encrypted, string attribute, Func<byte[], byte[]> change) =>
            RewriteDescriptor(encrypted, descriptor => {
                XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
                XAttribute value = descriptor.Root!.Element(ns + "dataIntegrity")!.Attribute(attribute)!;
                value.Value = Convert.ToBase64String(change(Convert.FromBase64String(value.Value)));
            });

        private static byte[] RewriteDescriptor(byte[] encrypted, Action<XDocument> change) {
            using var source = new MemoryStream();
            source.Write(encrypted, 0, encrypted.Length);
            source.Position = 0;
            using (RootStorage root = RootStorage.Open(source, StorageModeFlags.LeaveOpen)) {
                using CfbStream stream = root.OpenStream("EncryptionInfo");
                using var original = new MemoryStream();
                stream.CopyTo(original);
                byte[] bytes = original.ToArray();
                XDocument descriptor = XDocument.Parse(Encoding.UTF8.GetString(bytes, 8, bytes.Length - 8));
                change(descriptor);
                byte[] xml = Encoding.UTF8.GetBytes(descriptor.ToString(SaveOptions.DisableFormatting));
                stream.Position = 0;
                stream.SetLength(8 + xml.Length);
                stream.Write(bytes, 0, 8);
                stream.Write(xml, 0, xml.Length);
                root.Flush();
            }
            return source.ToArray();
        }

        private static byte[] RewritePayload(byte[] encrypted) {
            using var source = new MemoryStream();
            source.Write(encrypted, 0, encrypted.Length);
            source.Position = 0;
            using (RootStorage root = RootStorage.Open(source, StorageModeFlags.LeaveOpen)) {
                using CfbStream payload = root.OpenStream("EncryptedPackage");
                payload.Position = 9;
                int value = payload.ReadByte();
                payload.Position = 9;
                payload.WriteByte((byte)(value ^ 1));
                root.Flush();
            }
            return source.ToArray();
        }
    }
}
