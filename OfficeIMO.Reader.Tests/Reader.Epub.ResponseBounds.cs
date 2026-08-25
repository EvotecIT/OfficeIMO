using OfficeIMO.Epub;
#if NET8_0_OR_GREATER
using System.Buffers.Binary;
using System.Text;
#endif
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
#if NET8_0_OR_GREATER
    [Fact]
    public void EpubDocument_DoesNotPreallocateUntrustedDeclaredChapterLength() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-length-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            byte[] bytes = File.ReadAllBytes(epubPath);
            PatchCentralDirectoryUncompressedSize(bytes, "OEBPS/chapter1.xhtml", 128 * 1024 * 1024);

            long before = GC.GetAllocatedBytesForCurrentThread();
            using var stream = new MemoryStream(bytes, writable: false);
            try {
                _ = EpubDocument.Load(stream, new EpubReadOptions {
                    MaxChapterBytes = null,
                    MaxTotalUncompressedBytes = long.MaxValue
                });
            } catch (EpubReadException) {
                // ZIP implementations may reject the deliberately inconsistent entry.
            }
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

            Assert.True(allocated < 32L * 1024 * 1024, $"Unexpected allocation: {allocated} bytes.");
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    private static void PatchCentralDirectoryUncompressedSize(
        byte[] archive,
        string entryName,
        int declaredLength) {
        byte[] encodedName = Encoding.UTF8.GetBytes(entryName);
        for (int offset = 0; offset <= archive.Length - 46; offset++) {
            if (BinaryPrimitives.ReadUInt32LittleEndian(archive.AsSpan(offset, 4)) != 0x02014b50) continue;
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(archive.AsSpan(offset + 28, 2));
            if (nameLength != encodedName.Length || offset + 46 + nameLength > archive.Length) continue;
            if (!archive.AsSpan(offset + 46, nameLength).SequenceEqual(encodedName)) continue;

            BinaryPrimitives.WriteUInt32LittleEndian(
                archive.AsSpan(offset + 24, 4),
                checked((uint)declaredLength));
            return;
        }

        throw new InvalidOperationException($"Central-directory entry '{entryName}' was not found.");
    }
#endif
}
