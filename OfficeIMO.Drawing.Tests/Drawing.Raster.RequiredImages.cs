using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterRequiredImagesTests {
    [Theory]
    [InlineData("image")]
    [InlineData("group")]
    [InlineData("effect")]
    [InlineData("tile")]
    [InlineData("pattern")]
    public void RequiredImageDecodeIsSingleAndPropagatesThroughNestedScenes(string kind) {
        OfficeDrawing scene = Scene(kind);
        var codec = new ControlledCodec();
        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(scene, new OfficeDrawingRasterRenderOptions {
            ThrowOnImageDecodeFailure = true, ImageCodec = codec
        });
        Assert.Equal(1, codec.Calls);
        Assert.Equal(OfficeColor.Red, rendered.GetPixel(0, 0));
        Assert.Throws<NotSupportedException>(() => OfficeDrawingRasterRenderer.Render(scene,
            new OfficeDrawingRasterRenderOptions { ThrowOnImageDecodeFailure = true }));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RequiredImageDecodeRejectsFailureNullAndOversizedOutput(bool success, bool oversized) {
        var codec = new ControlledCodec { Success = success, ReturnNull = !oversized, Oversized = oversized };
        Assert.Throws<NotSupportedException>(() => OfficeDrawingRasterRenderer.Render(Scene("image"),
            new OfficeDrawingRasterRenderOptions {
                ThrowOnImageDecodeFailure = true, ImageCodec = codec, MaximumRasterPixels = 4
            }));
        Assert.Equal(1, codec.Calls);
    }

    [Fact]
    public void RequiredImageDecodeObservesCancellationAfterExternalCodecReturns() {
        using var cancellation = new CancellationTokenSource();
        var codec = new ControlledCodec { OnDecode = cancellation.Cancel };
        Assert.ThrowsAny<OperationCanceledException>(() => OfficeDrawingRasterRenderer.Render(Scene("effect"),
            new OfficeDrawingRasterRenderOptions {
                ThrowOnImageDecodeFailure = true, ImageCodec = codec, CancellationToken = cancellation.Token
            }));
    }

    private static OfficeDrawing Scene(string kind) {
        var image = new OfficeDrawing(2, 2);
        image.AddImage(new byte[] { 1, 2, 3 }, "application/test-raster",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 2, 2)));
        var scene = new OfficeDrawing(2, 2);
        if (kind == "pattern") scene.AddImagePattern(new byte[] { 1, 2, 3 }, "application/test-raster",
            new OfficeImagePatternLayout(new OfficeImagePlacement(0, 0, 2, 2), new OfficeImagePlacement(0, 0, 2, 2)));
        else if (kind == "group") scene.AddClippedDrawing(image, 0, 0, OfficeClipPath.Rectangle(2, 2));
        else if (kind == "effect") scene.AddEffectDrawing(image, OfficeTransform.Translate(0, 0));
        else if (kind == "tile") scene.AddTilingPattern(image, new OfficeImagePlacement(0, 0, 2, 2), 2, 2);
        else return image;
        return scene;
    }

    private sealed class ControlledCodec : IOfficeRasterImageCodec {
        internal int Calls;
        internal bool Success = true;
        internal bool ReturnNull;
        internal bool Oversized;
        internal Action? OnDecode;
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Assert.Equal(1, ++Calls);
            OnDecode?.Invoke();
            image = ReturnNull ? null : new OfficeRasterImage(Oversized ? 3 : 1, Oversized ? 3 : 1, OfficeColor.Red);
            return Success;
        }
    }
}
