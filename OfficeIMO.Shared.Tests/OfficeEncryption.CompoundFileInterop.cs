using System;
using System.IO;
using OfficeIMO.Core.Internal;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class EncryptedCompoundFileTests {
        [Theory]
        [InlineData(0)]
        [InlineData(32)]
        [InlineData(5003)]
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
