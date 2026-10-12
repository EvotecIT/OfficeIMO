using System;
using System.IO;
using OfficeIMO.Core.Internal;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class EncryptedCompoundFileTests {
        [Theory]
        [InlineData(32)]
        [InlineData(4096)]
        public void ByteBackedAndProviderStreamsWriteEquivalentContainers(int payloadSize) {
            byte[] info = new byte[400];
            byte[] payload = new byte[payloadSize];
            int infoOpens = 0;
            int payloadOpens = 0;
            byte[] memory = OfficeCompoundFileWriter.Write(new[] {
                new OfficeCompoundStream("EncryptionInfo", info, count: 37),
                new OfficeCompoundStream("EncryptedPackage", payload)
            });
            byte[] providers = OfficeCompoundFileWriter.Write(new[] {
                new OfficeCompoundStream("EncryptionInfo", 37, () => {
                    infoOpens++;
                    return new MemoryStream(info, 0, 37, writable: false);
                }),
                new OfficeCompoundStream("EncryptedPackage", payloadSize, () => {
                    payloadOpens++;
                    return new MemoryStream(payload, writable: false);
                })
            });

            Assert.Equal(memory, providers);
            Assert.Equal(1, infoOpens);
            Assert.Equal(1, payloadOpens);
            using var input = new MemoryStream(memory, writable: false);
            using RootStorage root = RootStorage.Open(input, StorageModeFlags.LeaveOpen);
            using CfbStream information = root.OpenStream("EncryptionInfo");
            using CfbStream package = root.OpenStream("EncryptedPackage");
            Assert.Equal(37, information.Length);
            Assert.Equal(payloadSize, package.Length);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(32)]
        [InlineData(4080)]
        [InlineData(4081)]
        [InlineData(5003)]
        [InlineData(8 * 1024 * 1024)]
        public void EncryptPackage_WritesIndependentlyReadableNamedStreams(int payloadSize) {
            byte[] plaintext = new byte[payloadSize];
            byte[] encrypted = OfficeEncryption.EncryptPackage(plaintext, "test-password",
                new OfficeEncryptionOptions { SpinCount = 1 });
            // An independent directory lookup must find both streams, whether
            // their data lives in the mini stream or ordinary sector chains.
            using var input = new MemoryStream(encrypted, writable: false);
            using RootStorage root = RootStorage.Open(input, StorageModeFlags.LeaveOpen);
            using CfbStream info = root.OpenStream("EncryptionInfo");
            using CfbStream payload = root.OpenStream("EncryptedPackage");
            Assert.Equal(4, info.ReadByte());
            Assert.Equal(0, info.ReadByte());
            Assert.Equal(4, info.ReadByte());
            Assert.Equal(0, info.ReadByte());
            using var reader = new BinaryReader(payload);
            Assert.Equal((ulong)payloadSize, reader.ReadUInt64());
            Assert.Equal(plaintext, OfficeEncryption.DecryptPackage(encrypted, "test-password"));
        }
    }
}
