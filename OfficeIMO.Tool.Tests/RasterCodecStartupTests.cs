using OfficeIMO.Drawing;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

[CollectionDefinition("Raster codec startup", DisableParallelization = true)]
public sealed class RasterCodecStartupCollection { }

[Collection("Raster codec startup")]
public sealed class RasterCodecStartupTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StartupLeavesCodecRegistrationToTheCaller(bool registerCodec) {
        IOfficeRasterImageCodec? previous = OfficeRasterImageCodecs.Platform;
        IOfficeRasterImageCodec? registered = registerCodec ? new CallerCodec() : null;
        OfficeRasterImageCodecs.Platform = registered;
        try {
            _ = new OfficeWorkflowRunner();
            Assert.Same(registered, OfficeRasterImageCodecs.Platform);
            using var output = new MemoryStream();
            using var error = new StringWriter();
            int exit = await OfficeImoToolApp.RunAsync(new[] { "help" }, Stream.Null, output, error);
            Assert.Equal(0, exit);
            Assert.Same(registered, OfficeRasterImageCodecs.Platform);
        } finally {
            OfficeRasterImageCodecs.Platform = previous;
        }
    }

    private sealed class CallerCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = null;
            return false;
        }
    }
}
