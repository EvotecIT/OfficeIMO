using OfficeIMO.Drawing;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

[CollectionDefinition("HEIC startup codec", DisableParallelization = true)]
public sealed class HeicStartupCodecCollection { }

[Collection("HEIC startup codec")]
public sealed class HeicStartupCodecTests {
    [Fact]
    public async Task AutomaticStartupPreservesTheRegisteredProcessCodec() {
        IOfficeRasterImageCodec? previous = OfficeRasterImageCodecs.Platform;
        var registered = new CallerCodec();
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
