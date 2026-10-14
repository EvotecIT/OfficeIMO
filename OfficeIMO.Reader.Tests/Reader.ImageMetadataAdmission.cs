using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Reader.Image;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderImageMetadataAdmissionTests {
    [Theory]
    [InlineData(ReaderProjection.Full, "direct")]
    [InlineData(ReaderProjection.TextAndMetadata, "direct")]
    [InlineData(ReaderProjection.Full, "lazy")]
    [InlineData(ReaderProjection.TextAndMetadata, "lazy")]
    [InlineData(ReaderProjection.Full, "batch")]
    [InlineData(ReaderProjection.TextAndMetadata, "batch")]
    [InlineData(ReaderProjection.Full, "folder")]
    [InlineData(ReaderProjection.TextAndMetadata, "folder")]
    public async Task MetadataInputLimitRemainsTerminalAcrossReaderRoutes(ReaderProjection projection, string route) {
        byte[] bytes = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue), OfficeImageExportFormat.Png);
        var native = new ReaderImageOptions { MetadataInspection = new OfficeImageMetadataInspectionOptions { MaximumInputBytes = bytes.Length - 1 } };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler(native).Build();
        var options = new ReaderOptions { Projection = projection, Inspection = new ReaderInspectionOptions() };
        string directory = Path.Combine(Path.GetTempPath(), "image-admission-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "source.png"); File.WriteAllBytes(path, bytes);
            async Task Read() {
                if (route == "direct") reader.ReadDocument(bytes, "source.png", options);
                else if (route == "lazy") reader.Read(bytes, "source.png", options).ToArray();
                else if (route == "batch") await reader.ReadDocumentsDetailedAsync(new[] { path }, options);
                else reader.ReadFolder(directory, options: options).ToArray();
            }
            ReaderResourceLimitException error = await Assert.ThrowsAsync<ReaderResourceLimitException>(Read);
            Assert.Equal("image-metadata.MaximumInputBytes", error.LimitName);
            Assert.Equal(bytes.Length - 1, error.Maximum);
            OfficeInspectionResourceLimitException owner = Assert.IsType<OfficeInspectionResourceLimitException>(error.InnerException);
            Assert.Equal("image-metadata", owner.Owner); Assert.Equal(bytes.Length - 1, owner.Maximum);
            // The metadata policy is opt-in and does not lower ordinary image admission.
            Assert.NotNull(reader.ReadDocument(bytes, "source.png", new ReaderOptions { Projection = projection }));
        } finally { Directory.Delete(directory, true); }
    }
}
