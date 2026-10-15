using OfficeIMO.Core.Internal;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderCompoundAdmissionLimitTests {
    [Theory]
    [InlineData("WordDocument", false)]
    [InlineData("WordDocument", true)]
    [InlineData("__properties_version1.0", false)]
    [InlineData("__properties_version1.0", true)]
    public async Task CompoundDetectionQuotaIsPublicAndTerminalWithRestoredPosition(string name, bool asynchronous) {
        var streams = new Dictionary<string, byte[]> { [name] = new byte[] { 1 } };
        var source = new OfficeCompoundFile(streams,
            new[] { new OfficeCompoundFileEntry(name, name, 2, 1) },
            new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5, 0));
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]>());
        using var input = new MemoryStream();
        input.WriteByte(42);
        input.Write(bytes, 0, bytes.Length);
        input.Position = 1;
        var reader = new OfficeDocumentReaderBuilder().Build();
        var options = new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent, MaxContainerEntries = 3 };
        ReaderResourceLimitException error = asynchronous
            ? await Assert.ThrowsAsync<ReaderResourceLimitException>(() => reader.DetectAsync(input, "source.bin", options))
            : Assert.Throws<ReaderResourceLimitException>(() => reader.Detect(input, "source.bin", options));
        Assert.Equal("Detection.MaxDirectoryEntries", error.LimitName);
        Assert.Equal(3, error.Maximum);
        Assert.NotNull(error.InnerException);
        Assert.Equal(typeof(OfficeDocumentReader).Assembly, error.InnerException!.GetType().Assembly);
        Assert.Equal(1, input.Position);
    }

    [Fact]
    public void MalformedCompoundStillHasUnknownDetectionRatherThanQuotaFailure() {
        byte[] bytes = new byte[512];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(bytes, 0);
        var reader = new OfficeDocumentReaderBuilder().Build();
        ReaderDetectionResult result = reader.Detect(bytes, "source.bin",
            new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent, MaxContainerEntries = 3 });
        Assert.Equal(ReaderInputKind.Unknown, result.ContentKind);
    }
}
