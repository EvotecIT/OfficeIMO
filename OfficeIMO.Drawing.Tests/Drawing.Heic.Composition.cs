using System;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

[CollectionDefinition("HEIC Core registration", DisableParallelization = true)]
public sealed class HeicCoreRegistrationCollection { }

[Collection("HEIC Core registration")]
public sealed class HeicCompositionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StrictRasterRenderingResolvesTheUnderlyingCodec(bool explicitCodec) {
        var previous=OfficeRasterImageCodecs.Platform;
        var codec=new SamplesCodec();
        try {
            OfficeRasterImageCodecs.Platform=codec;
            byte[] bytes=HeicTestFixture.Create(new[]{HeicTestFixture.Tile(40)});
            var scene=new OfficeDrawing(3,2);
            scene.AddImage(bytes,"image/heic",new OfficeImageProjection(new OfficeImagePlacement(0,0,3,2)));
            var rendered=OfficeDrawingRasterRenderer.Render(scene,new OfficeDrawingRasterRenderOptions {
                ThrowOnImageDecodeFailure=true,ImageCodec=explicitCodec ? codec : null
            });
            Assert.Equal((byte)40,rendered.GetPixel(0,0).R);
            Assert.Equal(1,codec.Calls);
        } finally {OfficeRasterImageCodecs.Platform=previous;}
    }

    [Theory]
    [InlineData(OfficeRasterCodecUnavailableReason.NoPlatformCodec,false)]
    [InlineData(OfficeRasterCodecUnavailableReason.WindowsHeifExtensionMissing,true)]
    [InlineData(OfficeRasterCodecUnavailableReason.WindowsHevcExtensionMissing,false)]
    public void StrictRasterRenderingRetainsCodecAvailability(OfficeRasterCodecUnavailableReason reason,bool explicitCodec) {
        var previous=OfficeRasterImageCodecs.Platform;
        var codec=new UnavailableCodec(reason);
        try {
            OfficeRasterImageCodecs.Platform=codec;
            var scene=new OfficeDrawing(3,2);
            scene.AddImage(HeicTestFixture.Create(new[]{HeicTestFixture.Tile(40)}),"image/heic",
                new OfficeImageProjection(new OfficeImagePlacement(0,0,3,2)));
            var error=Assert.Throws<OfficeRasterDecodeException>(()=>OfficeDrawingRasterRenderer.Render(scene,
                new OfficeDrawingRasterRenderOptions {ThrowOnImageDecodeFailure=true,ImageCodec=explicitCodec ? codec : null}));
            Assert.Equal(reason,error.DecodeInfo.CodecUnavailableReason);
            Assert.Equal("image-codec-unavailable",error.DecodeInfo.DiagnosticCode);
        } finally {OfficeRasterImageCodecs.Platform=previous;}
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RejectedHeifCannotBypassCompositionThroughRasterFallback(bool strict) {
        var codec=new SamplesCodec();
        byte[] bytes=HeicTestFixture.Create(new[]{HeicTestFixture.Tile(40,3,2,
            HeicTestFixture.Box("pasp",HeicTestFixture.Join(HeicTestFixture.U32(2),HeicTestFixture.U32(1))))});
        var scene=new OfficeDrawing(3,2);
        scene.AddImage(bytes,"image/heic",new OfficeImageProjection(new OfficeImagePlacement(0,0,3,2)));
        var options=new OfficeDrawingRasterRenderOptions {ImageCodec=codec,ThrowOnImageDecodeFailure=strict};
        if(strict) Assert.Throws<NotSupportedException>(()=>OfficeDrawingRasterRenderer.Render(scene,options));
        else Assert.Equal((byte)0,OfficeDrawingRasterRenderer.Render(scene,options).GetPixel(0,0).A);
        Assert.Equal(0,codec.Calls);
    }

    [Theory]
    [InlineData(1,40,3,2)]
    [InlineData(2,42,3,2)]
    [InlineData(3,45,3,2)]
    [InlineData(4,43,3,2)]
    [InlineData(5,40,2,3)]
    [InlineData(6,43,2,3)]
    [InlineData(7,45,2,3)]
    [InlineData(8,42,2,3)]
    public void ExifOrientationKeepsAuxiliaryAlphaOnTheSamePixels(int orientation,int first,int width,int height) {
        var metadata=new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Orientation,(ushort)orientation);
        var exif=new HeicTestFixture.Item("Exif",0,0,HeicTestFixture.Join(HeicTestFixture.U32(0),metadata.ExifProfile!));
        byte[] source=HeicTestFixture.Create(new[]{HeicTestFixture.Tile(40),HeicTestFixture.Tile(100,3,2,HeicTestFixture.Alpha()),exif},
            ("auxl",2,new[]{1}),("cdsc",3,new[]{1}));
        Assert.True(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions{ImageCodec=new SamplesCodec()},out var image,out var info),info.Diagnostic);
        Assert.Equal(width,image!.Width);Assert.Equal(height,image.Height);
        Assert.Equal((byte)first,image.GetPixel(0,0).R);
        for(int y=0;y<height;y++) for(int x=0;x<width;x++) {
            var pixel=image.GetPixel(x,y);
            Assert.Equal(pixel.R+60,pixel.A);
        }
    }

    [Fact]
    public void SignatureProbeRejectsExcessiveBrandTables() {
        byte[] brands=new byte[1024*1024];
        byte[] unknownBrand=HeicTestFixture.Text("zzzz");
        for(int p=0;p<brands.Length;p+=4) unknownBrand.CopyTo(brands,p);
        HeicTestFixture.Text("heic").CopyTo(brands,brands.Length-4);
        byte[] source=HeicTestFixture.Box("ftyp",HeicTestFixture.Join(HeicTestFixture.Text("zzzz"),HeicTestFixture.U32(0),brands));
        Assert.False(OfficeImageReader.HasHeicSignature(source,CancellationToken.None));
        using var cancelled=new CancellationTokenSource();cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(()=>OfficeImageReader.HasHeicSignature(source,cancelled.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExcessiveBrandTablesCannotReachAnUninspectedCodec(bool brandFirst) {
        byte[] brands=new byte[4097*4];
        for(int p=0;p<brands.Length;p+=4) HeicTestFixture.Text("zzzz").CopyTo(brands,p);
        HeicTestFixture.Text("heic").CopyTo(brands,brandFirst ? 0 : brands.Length-4);
        byte[] source=HeicTestFixture.Box("ftyp",HeicTestFixture.Join(HeicTestFixture.Text("zzzz"),HeicTestFixture.U32(0),brands));
        var codec=new SamplesCodec();
        var scene=new OfficeDrawing(3,2);
        scene.AddImage(source,"application/octet-stream",new OfficeImageProjection(new OfficeImagePlacement(0,0,3,2)));
        _=OfficeDrawingRasterRenderer.Render(scene,new OfficeDrawingRasterRenderOptions {ImageCodec=codec});
        Assert.Equal(0,codec.Calls);
        Assert.False(OfficeSvgImageRenderer.TryDecodeRasterForSvg(source,null,codec,CancellationToken.None,out _,out _));
        Assert.Equal(0,codec.Calls);
    }

    [Fact]
    public void DrawingCacheCannotHideChangesToPlatformCodecAvailability() {
        IOfficeRasterImageCodec? previous=OfficeRasterImageCodecs.Platform;
        try {
            OfficeRasterImageCodecs.Platform=new SamplesCodec();
            var diagnostics=new System.Collections.Generic.List<OfficeImageExportDiagnostic>();
            var options=new OfficeDrawingRasterRenderOptions {ImageCache=new OfficeRasterImageCache(1024*1024),DiagnosticSink=diagnostics};
            var drawing=new OfficeDrawing(3,2);
            drawing.AddImage(HeicTestFixture.Create(new[]{HeicTestFixture.Tile(40)}),"image/heic",
                new OfficeImageProjection(new OfficeImagePlacement(0,0,3,2)));
            Assert.Equal((byte)40,OfficeDrawingRasterRenderer.Render(drawing,options).GetPixel(0,0).R);
            OfficeRasterImageCodecs.Platform=null;
            Assert.Equal((byte)0,OfficeDrawingRasterRenderer.Render(drawing,options).GetPixel(0,0).A);
            Assert.Contains(diagnostics,d=>d.Code=="image-codec-unavailable");
        } finally {OfficeRasterImageCodecs.Platform=previous;}
    }

    [Fact]
    public void PlatformFallbackAndExplicitCodecHaveDistinctPrecedence() {
        IOfficeRasterImageCodec? previous=OfficeRasterImageCodecs.Platform;
        byte[] source=HeicTestFixture.Create(new[]{HeicTestFixture.Tile(10)});
        var platform=new SamplesCodec();
        try {
            OfficeRasterImageCodecs.Platform=platform;
            Assert.True(OfficeRasterImageDecoder.TryDecode(source,out _));
            var explicitCodec=new SamplesCodec();
            Assert.True(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions{ImageCodec=explicitCodec},out _,out _));
            Assert.False(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions{ImageCodec=new UnavailableCodec(OfficeRasterCodecUnavailableReason.NoPlatformCodec)},out _,out _));
            Assert.Equal(1,platform.Calls);Assert.Equal(1,explicitCodec.Calls);
            OfficeRasterImageCodecs.Platform=null;
            Assert.False(OfficeRasterImageDecoder.TryDecode(source,null,out _,out var info));
            Assert.Equal(OfficeRasterCodecUnavailableReason.NoPlatformCodec,info.CodecUnavailableReason);
        } finally {OfficeRasterImageCodecs.Platform=previous;}
    }

    [Fact]
    public void ToneMapUsesOnlyItsSdrBaseAndIgnoresTheGainMap() {
        byte[] source = HeicTestFixture.Create(new[] {
            new HeicTestFixture.Item("tmap",3,2,new byte[]{0}),
            HeicTestFixture.Tile(20),
            new HeicTestFixture.Item("hvc1",0,0,new byte[]{255})
        },("dimg",1,new[]{2,3}));
        var codec=new SamplesCodec();
        Assert.True(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions{ImageCodec=codec},out var image,out var info),info.Diagnostic);
        Assert.Equal((byte)20,image!.GetPixel(0,0).R);
        Assert.Equal(1,codec.Calls);
    }

    [Fact]
    public void SingleItemIsSanitizedBeforeCallingTheCodec() {
        byte[] source = HeicTestFixture.Create(new[] { HeicTestFixture.Tile(20, 3, 2, HeicTestFixture.Rotation(1), HeicTestFixture.Mirror(1)) });
        var codec = new SamplesCodec();
        Assert.True(OfficeRasterImageDecoder.TryDecode(source, new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out var info), info.Diagnostic);
        Assert.Equal(2, image!.Width); Assert.Equal(3, image.Height);
        // Rotate anticlockwise and then mirror left/right: the bottom-right sample moves to top-left.
        Assert.Equal((byte)25, image.GetPixel(0, 0).R);
        Assert.Equal(1, codec.Calls);
    }

    [Fact]
    public void GridUsesRowMajorTilesAndCropsPadding() {
        byte[] source = HeicTestFixture.Create(new[] {
            new HeicTestFixture.Item("grid", 5, 3, HeicTestFixture.Grid(2,2,5,3)),
            HeicTestFixture.Tile(10), HeicTestFixture.Tile(30), HeicTestFixture.Tile(50), HeicTestFixture.Tile(70)
        }, ("dimg",1,new[]{2,3,4,5}));
        var codec = new SamplesCodec();
        Assert.True(OfficeRasterImageDecoder.TryDecode(source, new OfficeRasterDecodeOptions { ImageCodec=codec }, out var image, out var info), info.Diagnostic);
        Assert.Equal(5,image!.Width); Assert.Equal(3,image.Height);
        Assert.Equal((byte)10,image.GetPixel(0,0).R); Assert.Equal((byte)31,image.GetPixel(4,0).R);
        Assert.Equal((byte)50,image.GetPixel(0,2).R); Assert.Equal((byte)71,image.GetPixel(4,2).R);
        Assert.Equal(4,codec.Calls);
    }

    [Fact]
    public void CleanApertureRunsBeforeRotation() {
        byte[] source = HeicTestFixture.Create(new[] { HeicTestFixture.Tile(10,4,2,HeicTestFixture.Crop(2,2),HeicTestFixture.Rotation(1)) });
        Assert.True(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions { ImageCodec=new SamplesCodec() },out var image,out var info),info.Diagnostic);
        Assert.Equal(2,image!.Width); Assert.Equal(2,image.Height);
        Assert.Equal((byte)12,image.GetPixel(0,0).R); Assert.Equal((byte)15,image.GetPixel(1,1).R);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void AlphaUsesAuxlAndUnpremultipliesBeforeColorConversion(bool premultiplied) {
        var items = new[] { HeicTestFixture.Tile(40),HeicTestFixture.Tile(128,3,2,HeicTestFixture.Alpha()) };
        byte[] source = premultiplied ? HeicTestFixture.Create(items,("auxl",2,new[]{1}),("prem",1,new[]{2})) : HeicTestFixture.Create(items,("auxl",2,new[]{1}));
        Assert.True(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions { ImageCodec=new SamplesCodec() },out var image,out var info),info.Diagnostic);
        Assert.Equal((byte)128,image!.GetPixel(0,0).A);
        Assert.Equal(premultiplied ? (byte)80 : (byte)40,image.GetPixel(0,0).R);
    }

    [Theory]
    [InlineData(false,2,3)]
    [InlineData(true,3,2)]
    public void ContainerOrientationTakesPrecedenceOverExif(bool hasTransform,int width,int height) {
        var metadata=new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Orientation,(ushort)6);
        var exif=new HeicTestFixture.Item("Exif",0,0,HeicTestFixture.Join(HeicTestFixture.U32(0),metadata.ExifProfile!));
        var tile=hasTransform ? HeicTestFixture.Tile(10,3,2,HeicTestFixture.Rotation(0)) : HeicTestFixture.Tile(10);
        byte[] source=HeicTestFixture.Create(new[]{tile,exif},("cdsc",2,new[]{1}));
        Assert.True(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions{ImageCodec=new SamplesCodec()},out var image,out var info),info.Diagnostic);
        Assert.Equal(width,image!.Width);Assert.Equal(height,image.Height);
    }

    [Theory]
    [InlineData(OfficeRasterCodecUnavailableReason.NoPlatformCodec)]
    [InlineData(OfficeRasterCodecUnavailableReason.WindowsHeifExtensionMissing)]
    [InlineData(OfficeRasterCodecUnavailableReason.WindowsHevcExtensionMissing)]
    public void MissingPlatformComponentsKeepTheirDiagnostic(OfficeRasterCodecUnavailableReason reason) {
        byte[] source=HeicTestFixture.Create(new[]{HeicTestFixture.Tile(1)});
        Assert.False(OfficeRasterImageDecoder.TryDecode(source,new OfficeRasterDecodeOptions{ImageCodec=new UnavailableCodec(reason)},out _,out var info));
        Assert.Equal("image-codec-unavailable",info.DiagnosticCode);
        Assert.Equal(reason,info.CodecUnavailableReason);
        Assert.Equal(OfficeRasterDecodeFailure.CodecUnavailable,info.Failure);
    }

    private sealed class UnavailableCodec : IOfficeRasterImageCodec, IOfficeRasterImageCodecAvailability {
        private readonly OfficeRasterCodecUnavailableReason _reason;
        internal UnavailableCodec(OfficeRasterCodecUnavailableReason reason)=>_reason=reason;
        public OfficeRasterCodecAvailability GetAvailability()=>new OfficeRasterCodecAvailability(_reason,"Install the required system image component.");
        public bool TryDecode(byte[] bytes,string? type,out OfficeRasterImage? image) {image=null;throw new InvalidOperationException("An unavailable codec must not decode.");}
    }

    [Fact]
    public void InvalidGridAndBudgetRejectionNeverCallTheCodec() {
        var codec=new SamplesCodec();
        byte[] invalid=HeicTestFixture.Create(new[] {new HeicTestFixture.Item("grid",5,3,HeicTestFixture.Grid(2,2,5,3)),HeicTestFixture.Tile(1)},("dimg",1,new[]{2}));
        Assert.False(OfficeRasterImageDecoder.TryDecode(invalid,new OfficeRasterDecodeOptions{ImageCodec=codec},out _,out _));
        byte[] tile=HeicTestFixture.Create(new[] {HeicTestFixture.Tile(1)});
        Assert.False(OfficeRasterImageDecoder.TryDecode(tile,new OfficeRasterDecodeOptions{ImageCodec=codec,MaximumDecodedPixels=5},out _,out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(tile,new OfficeRasterDecodeOptions{ImageCodec=codec,MaximumEncodedBytes=tile.Length-1},out _,out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(tile,new OfficeRasterDecodeOptions{ImageCodec=codec,RetainedManagedBytes=256L*1024*1024-1},out _,out _));
        using var cancelled=new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(()=>OfficeRasterImageDecoder.TryDecode(tile,new OfficeRasterDecodeOptions{ImageCodec=codec,CancellationToken=cancelled.Token},out _,out _));
        Assert.Equal(0,codec.Calls);
    }

    [Theory]
    [InlineData("basic-8bit-srgb.heic",1600,1200)]
    [InlineData("generic-mif1-brand.heif",1600,1200)]
    [InlineData("grid-12mp-512px-tiles.heic",4032,3024)]
    [InlineData("orientation-6.heic",1200,800)]
    public void PublicSamplesAreIdentifiedWithoutDecoding(string name,int width,int height) {
        byte[] bytes=File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),"OfficeIMO.TestAssets","Images","Heic",name));
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes,null,out var image));
        Assert.Equal(OfficeImageFormat.Heic,image.Format); Assert.Equal(width,image.Width);Assert.Equal(height,image.Height);
    }

    private sealed class SamplesCodec : IOfficeRasterImageCodec {
        internal int Calls {get;private set;}
        public bool TryDecode(byte[] bytes,string? contentType,out OfficeRasterImage? image) {
            Calls++;
            Assert.Equal("image/heic",contentType);
            Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes,out var info));
            var item=Assert.Single(info!.Items);
            Assert.Equal("hvc1",item.ItemType);
            Assert.DoesNotContain(item.PropertyAssociations,p=>p.PropertyType is "colr" or "irot" or "imir" or "clap");
            int marker=bytes[item.Location!.Extents[0].Offset];
            image=new OfficeRasterImage((int)item.Width!.Value,(int)item.Height!.Value);
            for(int y=0;y<image.Height;y++) for(int x=0;x<image.Width;x++) image.SetPixel(x,y,OfficeColor.FromRgb((byte)(marker+y*image.Width+x),0,0));
            return true;
        }
    }
}
