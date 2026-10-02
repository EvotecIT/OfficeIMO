using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Public byte/stream contracts protect default AVIF decode and bounded caller-codec behavior.</summary>
public sealed class DrawingAvifPublicDecodeTests {
    [Theory]
    [InlineData("avif-opaque")]
    [InlineData("avif-alpha")]
    [InlineData("avif-monochrome-full")]
    [InlineData("avif-monochrome-limited")]
    [InlineData("avif-monochrome-full-alpha")]
    [InlineData("avif-monochrome-limited-alpha")]
    [InlineData("avif-main10-420-full")]
    [InlineData("avif-main10-420-limited")]
    [InlineData("avif-main10-420-full-alpha")]
    [InlineData("avif-main10-420-limited-alpha")]
    [InlineData("avif-main10-mono-full")]
    [InlineData("avif-main10-mono-limited")]
    [InlineData("avif-main10-mono-full-alpha")]
    [InlineData("avif-main10-mono-limited-alpha")]
    public void DefaultByteAndStreamDecodePreservesIndependentPixelsAndSelection(string name) {
        byte[] bytes=File.ReadAllBytes(Asset(name+".avif")),saved=(byte[])bytes.Clone(),expected=File.ReadAllBytes(Asset(name+".rgba"));
        var codec=new RejectingCodec();var options=new OfficeRasterDecodeOptions {ImageCodec=codec};
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes,options,out var image,out var info));
        Assert.Equal(OfficeImageFormat.Avif,info.Format);Assert.Equal(1,info.FrameCount);Assert.True(info.Succeeded);
        Assert.False(info.AnimationDiscarded);Assert.False(info.FramesOrPagesDiscarded);Assert.Null(info.Diagnostic);
        Assert.Equal("image/avif",OfficeImageInfo.GetMimeType(info.Format));Assert.Equal(".avif",OfficeImageInfo.GetDefaultExtension(info.Format));
        Assert.Equal(OfficeImageFormat.Avif,OfficeImageInfo.FromMimeType("IMAGE/AVIF; charset=binary"));
        Assert.Equal(OfficeImageFormat.Avif,OfficeImageReader.FromExtension(name+".avif"));
        Assert.True(OfficeImageReader.TryValidateContent(bytes,name+".avif",out var metadata));Assert.Equal(49,metadata.Width);Assert.Equal(33,metadata.Height);
        Assert.True(OfficeRasterContainerInspector.TryInspect(bytes,out var inspected));Assert.Equal(1,inspected!.Count);
        Assert.Equal(expected.Length,image!.PixelBuffer.Length);
        int rgbTolerance = name.StartsWith("avif-monochrome-", StringComparison.Ordinal) || name.StartsWith("avif-main10-", StringComparison.Ordinal) ? 1 : 3;
        for(int i=0;i<expected.Length;i++)Assert.InRange(Math.Abs(expected[i]-image.PixelBuffer[i]),0,i%4==3?0:rgbTolerance);
        using var stream=new MemoryStream(bytes);Assert.True(OfficeRasterImageDecoder.TryDecode(stream,options,out var streamImage,out _));
        Assert.Equal(0,stream.Position);Assert.Equal(image.PixelBuffer,streamImage!.PixelBuffer);
        Assert.Equal(saved,bytes);Assert.Equal(0,codec.Calls);
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,new OfficeRasterDecodeOptions {FrameIndex=1},out var missing,out var selection));
        Assert.Null(missing);Assert.Equal(1,selection.FrameCount);
    }

    [Theory]
    [InlineData("avif-alpha")]
    [InlineData("avif-monochrome-limited-alpha")]
    [InlineData("avif-main10-420-full-alpha")]
    [InlineData("avif-main10-mono-limited-alpha")]
    public void PixelEncodedRetainedAndCanceledRequestsCannotReachTheCallerCodec(string name) {
        byte[] bytes=File.ReadAllBytes(Asset(name + ".avif"));var codec=new RejectingCodec();
        foreach(var options in new[] {
            new OfficeRasterDecodeOptions {MaximumDecodedPixels=49*33-1},
            new OfficeRasterDecodeOptions {MaximumEncodedBytes=bytes.Length-1},
            new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes}
        }) {
            options.ImageCodec=codec;Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,options,out var image,out _));Assert.Null(image);
        }
        using var cancellation=new CancellationTokenSource();cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>OfficeRasterImageDecoder.TryDecode(bytes,new OfficeRasterDecodeOptions {
            CancellationToken=cancellation.Token,ImageCodec=codec},out _,out _));Assert.Equal(0,codec.Calls);
    }

    [Theory]
    [InlineData("avif-alpha")]
    [InlineData("avif-main10-420-full-alpha")]
    [InlineData("avif-main10-mono-limited-alpha")]
    public void MalformedAlphaPayloadCannotPublishColorOnlyAndInspectedUnsupportedColorMayUseTheTrustedCodec(string name) {
        byte[] original=File.ReadAllBytes(Asset(name+".avif")),broken=(byte[])original.Clone();
        Assert.True(OfficeAvifContainerReader.TryRead(original,new OfficeRasterDecodeOptions(),out var container));
        broken[container!.Alpha!.Offset]=0x80; // forbidden OBU bit in the selected alpha item
        Assert.False(OfficeRasterImageDecoder.TryDecode(broken,out var image));Assert.Null(image);
        Assert.False(OfficeImageReader.TryValidateContent(broken,"image.avif",out _));
        byte[] unsupported=(byte[])original.Clone();int nclx=Find(unsupported,"nclx");
        unsupported[nclx+8]=0;unsupported[nclx+9]=10; // valid metadata, unsupported constant-luminance matrix
        var codec=new RejectingCodec();Assert.False(OfficeRasterImageDecoder.TryDecode(unsupported,new OfficeRasterDecodeOptions {ImageCodec=codec},out image,out var info));
        Assert.Null(image);Assert.Equal(OfficeImageFormat.Avif,info.Format);Assert.Equal(1,codec.Calls);Assert.Equal("image/avif",codec.Mime);
        var accepting = new AcceptingCodec();
        Assert.True(OfficeRasterImageDecoder.TryDecode(unsupported, new OfficeRasterDecodeOptions { ImageCodec = accepting }, out image, out _));
        Assert.Equal(1, accepting.Calls); Assert.Equal(49, image!.Width); Assert.Equal(33, image.Height);
        Assert.Equal(original,File.ReadAllBytes(Asset(name+".avif")));
    }

    [Fact]
    public void PostInspectionAvifSafetyFailuresCannotBecomeCallerCodecSuccess() {
        byte[] bytes = File.ReadAllBytes(Asset("avif-alpha.avif"));
        var codec = new AcceptingCodec();
        foreach (var options in new[] {
            new OfficeRasterDecodeOptions { RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - 100000 },
            new OfficeRasterDecodeOptions { MaximumInspectionWorkPixels = 49 * 33 - 1 }
        }) {
            options.ImageCodec = codec;
            Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, options, out var image, out _));
            Assert.Null(image);
        }
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out var container));
        bytes[container!.Alpha!.Offset] = 0x80;
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { ImageCodec = codec }, out var broken, out _));
        Assert.Null(broken);
        Assert.Equal(0, codec.Calls);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void TruncatedAvifCannotUseSvgCustomFormatFallback(bool compatibleBrand, bool leadingFreeBox) {
        byte[] bytes = File.ReadAllBytes(Asset("avif-alpha.avif"));
        Array.Resize(ref bytes, bytes.Length - 1);
        if (compatibleBrand) Buffer.BlockCopy(Encoding.ASCII.GetBytes("mif1"), 0, bytes, 8, 4);
        if (leadingFreeBox) {
            byte[] prefixed = new byte[bytes.Length + 8];
            prefixed[3] = 8; Buffer.BlockCopy(Encoding.ASCII.GetBytes("free"), 0, prefixed, 4, 4);
            Buffer.BlockCopy(bytes, 0, prefixed, 8, bytes.Length); bytes = prefixed;
        }
        var codec = new AcceptingCodec();
        var drawing = new OfficeDrawing(49, 33).AddImageWithInterpolation(bytes, "image/avif",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 49, 33)), interpolate: false);
        Assert.Throws<InvalidOperationException>(() => OfficeDrawingSvgExporter.ToSvg(
            drawing, 1D, OfficeSvgSizeUnit.Pixel, imageCodec: codec));
        Assert.Equal(0, codec.Calls);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RejectedAvifCannotBypassSafetyThroughRasterRenderer(bool truncated) {
        byte[] bytes = File.ReadAllBytes(Asset("avif-alpha.avif"));
        if (truncated) Array.Resize(ref bytes, bytes.Length - 1);
        else {
            Assert.True(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out var container));
            bytes[container!.Alpha!.Offset] = 0x80;
        }
        var codec = new AcceptingCodec();
        var drawing = new OfficeDrawing(49, 33).AddImageWithInterpolation(bytes, "image/avif",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 49, 33)), interpolate: false);
        OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions { ImageCodec = codec });
        Assert.Equal(0, codec.Calls);
    }

    private static int Find(byte[] bytes,string text) {
        byte[] marker=Encoding.ASCII.GetBytes(text);
        for(int i=0;i<=bytes.Length-marker.Length;i++)if(bytes.Skip(i).Take(marker.Length).SequenceEqual(marker))return i;
        throw new InvalidOperationException("Missing fixture marker.");
    }
    private static string Asset(string name)=>Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",name);
    private sealed class RejectingCodec:IOfficeRasterImageCodec {
        internal int Calls;internal string? Mime;
        public bool TryDecode(byte[] bytes,string? contentType,out OfficeRasterImage? image) {Calls++;Mime=contentType;image=null;return false;}
    }
    private sealed class AcceptingCodec : IOfficeRasterImageCodec {
        internal int Calls;
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++; image = new OfficeRasterImage(49, 33, OfficeColor.Red); return true;
        }
    }
}
