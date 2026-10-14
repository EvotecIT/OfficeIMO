using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public partial class VisioForeignImageRenderingTests {
    // Independent two-frame 16x16 WebP from the shared Drawing codec fixture.
    private const string AnimatedWebp = "UklGRogAAABXRUJQVlA4WAoAAAACAAAADwAADwAAQU5JTQYAAAAAAAAAAABBTk1GKgAAAAAAAAAAAA8AAA8AAGQAAAJWUDhMEQAAAC8PwAMAB1CoohSv/4GI6H8AAEFOTUYqAAAAAAAAAAAADwAADwAAZAAAAFZQOEwRAAAALw/AAwAHUKjiFaX/gYjofwAA";

    [Fact]
    public void CallerDecodedAnimationReportsLossAndRejectsStrictExport() {
        var document = WithImage(Convert.FromBase64String(AnimatedWebp));
        var codec = new SuccessfulCodec(16);
        var options = new OfficeIMO.Visio.VisioImageExportOptions { RenderText = false, ImageCodec = codec };
        var result = document.ExportImage(OfficeImageExportFormat.Svg, options);
        Assert.Equal(1, codec.Calls);
        Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageStaticFrameSelected && d.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
        options.Policy.RequireNoOmissions = true;
        Assert.Throws<OfficeImageExportPolicyException>(() => document.ExportImage(OfficeImageExportFormat.Png, options));
    }

    [Fact]
    public void CallerCannotReplaceInspectedRasterDimensions() {
        var document = WithImage(Convert.FromBase64String(AnimatedWebp));
        var codec = new SuccessfulCodec(1);
        var result = document.ExportImage(OfficeImageExportFormat.Svg, new OfficeIMO.Visio.VisioImageExportOptions { ImageCodec = codec, RenderText = false });
        Assert.Equal(1, codec.Calls);
        Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback && d.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void SourceRasterLimitIsEnforcedBeforeCallerCodec() {
        byte[] gif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
        gif[6] = gif[8] = 0x70; gif[7] = gif[9] = 0x17; // 6000 x 6000 canvas.
        Assert.True(OfficeImageReader.TryIdentifyByContent(gif, null, out var info));
        Assert.Equal(6000, info.Width);
        var codec = new SuccessfulCodec(16);
        var result = WithImage(gif).ExportImage(OfficeImageExportFormat.Svg, new VisioImageExportOptions { ImageCodec = codec, RenderText = false });
        Assert.Equal(0, codec.Calls);
        Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
    }

    [Fact]
    public void IdentifiedMetafileUsesExplicitCallerCodecBoundary() {
        // Standard WMF header, one SetBkColor record and an EOF record.
        byte[] wmf = { 1,0,9,0,0,3,17,0,0,0,0,0,5,0,0,0,0,0,5,0,0,0,1,2,0,0,0,0,3,0,0,0,0,0 };
        var document = WithImage(wmf, "Metafile");
        var codec = new SuccessfulCodec(16);
        var result = document.ExportImage(OfficeImageExportFormat.Svg, new OfficeIMO.Visio.VisioImageExportOptions { ImageCodec = codec, RenderText = false });
        Assert.Equal(1, codec.Calls);
        Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
        Assert.DoesNotContain(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
    }

    private static OfficeIMO.Visio.VisioDocument WithImage(byte[] bytes, string type = "Bitmap") {
        using var input = Fixture(); var xml = XDocument.Load(input);
        var foreign = xml.Descendants(Ns + "ForeignData").Single();
        foreign.SetAttributeValue("ForeignType", type); foreign.SetAttributeValue("CompressionType", null);
        foreign.Value = Convert.ToBase64String(bytes);
        return Load(xml);
    }
    private sealed class SuccessfulCodec(int size) : IOfficeRasterImageCodec {
        public int Calls { get; private set; }
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++; image = new OfficeRasterImage(size, size, OfficeColor.Red); return true;
        }
    }
}
