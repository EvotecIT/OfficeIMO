using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class PublicImageMetadataTests {
    [Fact]
    public void SharedTiffStripArraysAreRejectedBeforeAggregateReferenceExpansion() {
        const int count = 65535;
        const int tableBytes = 54;
        int offsets = 8 + 2 * tableBytes;
        int lengths = offsets + count * 4;
        int pixel = lengths + count * 4;
        byte[] encoded = new byte[pixel + 1]; encoded[0] = 73; encoded[1] = 73; encoded[2] = 42; encoded[4] = 8;
        for (int page = 0; page < 2; page++) {
            int table = 8 + page * tableBytes; Put(encoded, table, 4, 2, true);
            Entry(table + 2, 256, 1, 1); Entry(table + 14, 257, 1, count); Entry(table + 26, 273, count, offsets); Entry(table + 38, 279, count, lengths);
            Put(encoded, table + 50, page == 0 ? 62U : 0U, 4, true);
        }
        for (int index = 0; index < count; index++) { Put(encoded, offsets + index * 4, (uint)pixel, 4, true); Put(encoded, lengths + index * 4, 1, 4, true); }
        var metadata = new OfficeImageMetadata(); metadata.SetExifValue(OfficeExifTag.Artist, "Edited");
        Assert.True(OfficeImageReader.TryIdentifyByContent(encoded, null, System.Threading.CancellationToken.None, out _));
        Assert.True(OfficeTiffStructureValidator.TryValidate(encoded, 0, encoded.Length));
        FormatException apply = Assert.Throws<FormatException>(() => OfficeImageMetadata.Apply(encoded, metadata));
        Assert.Contains("aggregate pixel-reference", apply.Message);
        FormatException remove = Assert.Throws<FormatException>(() => OfficeImageMetadata.Remove(encoded, OfficeImageMetadataProfileKinds.Exif));
        Assert.Contains("aggregate pixel-reference", remove.Message);
        Assert.Equal(1U, Read(encoded, lengths, true));
        void Entry(int at, uint tag, uint elements, int value) { Put(encoded, at, tag, 2, true); Put(encoded, at + 2, 4, 2, true); Put(encoded, at + 4, elements, 4, true); Put(encoded, at + 8, (uint)value, 4, true); }
    }

    [Theory]
    [InlineData(273, 279, false)]
    [InlineData(324, 325, false)]
    [InlineData(513, 514, false)]
    [InlineData(324, 325, true)]
    public void TiffAliasProtectionIncludesAllPixelFamiliesAndChildDirectories(int offsetTag, int lengthTag, bool childDirectory) {
        byte[] tiff = OfficeTiffCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Blue));
        var metadata = OfficeImageMetadata.Read(tiff); metadata.SetExifValue(OfficeExifTag.Artist, "Private creator");
        tiff = OfficeImageMetadata.Apply(tiff, metadata);
        int root = (int)Read(tiff, 4, true); int count = tiff[root] | tiff[root + 1] << 8;
        int artistOffset = 0, artistLength = 0;
        for (int index = 0; index < count; index++) { int at = root + 2 + index * 12; if ((tiff[at] | tiff[at + 1] << 8) == 315) { artistLength = (int)Read(tiff, at + 4, true); artistOffset = (int)Read(tiff, at + 8, true); } }
        int additions = childDirectory ? 1 : offsetTag == 273 ? 0 : 2;
        int newRoot = tiff.Length; int newCount = count + additions; int child = newRoot + 6 + newCount * 12;
        var annotated = new byte[child + (childDirectory ? 30 : 0)]; Buffer.BlockCopy(tiff, 0, annotated, 0, tiff.Length);
        Put(annotated, 4, (uint)newRoot, 4, true); Put(annotated, newRoot, (uint)newCount, 2, true); Buffer.BlockCopy(tiff, root + 2, annotated, newRoot + 2, count * 12);
        if (childDirectory) { PixelEntry(newRoot + 2 + count * 12, 330, child); Put(annotated, child, 2, 2, true); PixelEntry(child + 2, offsetTag, artistOffset); PixelEntry(child + 14, lengthTag, artistLength); }
        else if (offsetTag != 273) { PixelEntry(newRoot + 2 + count * 12, offsetTag, artistOffset); PixelEntry(newRoot + 14 + count * 12, lengthTag, artistLength); }
        else for (int index = 0; index < count; index++) { int at = newRoot + 2 + index * 12; int tag = annotated[at] | annotated[at + 1] << 8; if (tag == offsetTag) Put(annotated, at + 8, (uint)artistOffset, 4, true); if (tag == lengthTag) Put(annotated, at + 8, (uint)artistLength, 4, true); }
        OfficeImageMetadata edited = OfficeImageMetadata.Read(annotated); edited.RemoveExifValue(OfficeExifTag.Artist);
        Assert.Contains("overlaps encoded pixel", Assert.Throws<FormatException>(() => OfficeImageMetadata.Apply(annotated, edited)).Message);
        Assert.Contains("Private creator", Encoding.ASCII.GetString(annotated));
        void PixelEntry(int at, int tag, int value) { Put(annotated, at, (uint)tag, 2, true); Put(annotated, at + 2, 4, 2, true); Put(annotated, at + 4, 1, 4, true); Put(annotated, at + 8, (uint)value, 4, true); }
    }

    [Fact]
    public void InternalContainerExifParseAndExportChargeRetainedCallerBytes() {
        var metadata = new OfficeImageMetadata(); metadata.SetExifValue(OfficeExifTag.Artist, "Independent"); byte[] profile = metadata.EncodeExifProfile()!;
        Assert.Throws<FormatException>(() => OfficeImageMetadata.ParseExifProfile(profile, OfficeRasterGuards.MaximumDecodedBytes, System.Threading.CancellationToken.None));
        OfficeImageMetadata parsed = OfficeImageMetadata.ParseExifProfile(profile, 128L * 1024L * 1024L, System.Threading.CancellationToken.None);
        Array.Clear(profile, 0, profile.Length);
        Assert.Equal("Independent", parsed.GetExifValue(OfficeExifTag.Artist)!.Value);
        Assert.NotNull(parsed.EncodeExifProfile());
        Assert.Throws<ArgumentException>(() => parsed.EncodeExifProfile(System.Threading.CancellationToken.None, OfficeRasterGuards.MaximumDecodedBytes));
        using var token = new System.Threading.CancellationTokenSource(); token.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeTiffPixelRanges.Read(OfficeTiffCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue)), token.Token));
    }
    [Fact]
    public void JpegDensityEditsCreateMissingCarrierAndSynchronizeExistingExif() {
        var options = new OfficeRasterEncodingOptions();
        options.Jpeg.WriteJfifHeader = false;
        byte[] original = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue), OfficeImageExportFormat.Jpeg, options);
        var metadata = OfficeImageMetadata.Read(original);
        metadata.HorizontalResolution = 60;
        metadata.VerticalResolution = 50;
        metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerCentimeter;
        byte[] jfif = OfficeImageMetadata.Apply(original, metadata);
        OfficeImageMetadata read = OfficeImageMetadata.Read(jfif);
        Assert.Equal(60D, read.HorizontalResolution);
        Assert.Equal(50D, read.VerticalResolution);
        Assert.Equal(OfficeImageResolutionUnit.PixelsPerCentimeter, read.ResolutionUnits);
        Assert.Equal(JpegScan(original), JpegScan(jfif));
        Assert.Null(read.ExifProfile);
        read.SetExifValue(OfficeExifTag.Software, "Density test");
        read.SetExifValue(OfficeExifTag.XResolution, new OfficeRational(96, 1));
        read.SetExifValue(OfficeExifTag.YResolution, new OfficeRational(96, 1));
        read.SetExifValue(OfficeExifTag.ResolutionUnit, (ushort)2);
        read.HorizontalResolution = 80;
        read.VerticalResolution = 70;
        OfficeImageMetadata synchronized = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(jfif, read));
        Assert.Equal(80D, synchronized.HorizontalResolution);
        Assert.Equal(80D, Assert.IsType<OfficeRational>(synchronized.GetExifValue(OfficeExifTag.XResolution)!.Value).ToDouble());
        Assert.Equal(70D, Assert.IsType<OfficeRational>(synchronized.GetExifValue(OfficeExifTag.YResolution)!.Value).ToDouble());
        Assert.Equal((ushort)3, synchronized.GetExifValue(OfficeExifTag.ResolutionUnit)!.Value);
    }

    [Fact]
    public void GifAspectRatioIsEditedWithoutChangingImageBlocks() {
        byte[] gif = Convert.FromBase64String("R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=");
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(gif);
        Assert.Equal(OfficeImageResolutionUnit.AspectRatio, metadata.ResolutionUnits);
        metadata.HorizontalResolution = 2;
        metadata.VerticalResolution = 1;
        byte[] output = OfficeImageMetadata.Apply(gif, metadata);
        Assert.Equal(113, output[12]);
        Assert.Equal(SliceBytes(gif, 13, gif.Length - 13), SliceBytes(output, 13, output.Length - 13));
        Assert.Equal(2D, OfficeImageMetadata.Read(output).HorizontalResolution);
        metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerInch;
        Assert.Throws<NotSupportedException>(() => OfficeImageMetadata.Apply(gif, metadata));
        OfficeImageMetadata projected = metadata.PrepareForEncoding(OfficeImageFormat.Gif, out _);
        Assert.Equal(OfficeImageResolutionUnit.AspectRatio, projected.ResolutionUnits);
        Assert.Equal(OfficeImageResolutionUnit.PixelsPerInch, metadata.ResolutionUnits);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Webp)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Bmp)]
    public void MalformedIccCannotBeWrittenToSupportedContainers(OfficeImageExportFormat format) {
        byte[] image = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(2, 1, OfficeColor.Red), format);
        var metadata = OfficeImageMetadata.Read(image);
        metadata.IccProfile = new byte[] { 1, 2, 3 };
        Assert.Throws<FormatException>(() => OfficeImageMetadata.Apply(image, metadata));
        Assert.True(OfficeRasterImageDecoder.TryDecode(image, out _));
    }

    [Theory]
    [InlineData(OfficeImageResolutionUnit.AspectRatio, 4D, 3D)]
    [InlineData(OfficeImageResolutionUnit.PixelsPerCentimeter, 60D, 50D)]
    public void NativeTiffResolutionIsWrittenToEveryPageAndStreamingOutput(OfficeImageResolutionUnit unit, double x, double y) {
        var options = new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(x, y, unit) };
        var pages = new[] { new OfficeRasterImage(2, 2, OfficeColor.Red), new OfficeRasterImage(3, 1, OfficeColor.Blue) };
        byte[] encoded = OfficeTiffCodec.EncodePages(pages, options);
        Assert.True(OfficeRasterContainerInspector.TryInspect(encoded, out OfficeRasterContainerInfo? info));
        Assert.Equal(2, info!.Count);
        foreach (OfficeRasterFrameInfo page in info.Frames) {
            if (unit == OfficeImageResolutionUnit.AspectRatio) { Assert.Null(page.DpiX); Assert.Null(page.DpiY); }
            else { Assert.Equal(x * 2.54D, page.DpiX!.Value, 6); Assert.Equal(y * 2.54D, page.DpiY!.Value, 6); }
        }
        bool little = encoded[0] == 73;
        int at = (int)Read(encoded, 4, little);
        for (int index = 0; index < 2; index++) {
            var pageView = (byte[])encoded.Clone(); Put(pageView, 4, (uint)at, 4, little);
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(pageView);
            Assert.Equal(unit, metadata.ResolutionUnits); Assert.Equal(x, metadata.HorizontalResolution); Assert.Equal(y, metadata.VerticalResolution);
            int count = pageView[at] | pageView[at + 1] << 8; at = (int)Read(encoded, at + 2 + count * 12, little);
        }
        using var destination = new MemoryStream(); OfficeTiffCodec.EncodeTo(pages[0], destination, options);
        OfficeImageMetadata streamed = OfficeImageMetadata.Read(destination.ToArray());
        Assert.Equal(unit, streamed.ResolutionUnits); Assert.Equal(x, streamed.HorizontalResolution);
        var generic = new OfficeRasterEncodingOptions { Tiff = options };
        OfficeImageMetadata routed = OfficeImageMetadata.Read(OfficeRasterImageEncoder.Encode(pages[0], OfficeImageExportFormat.Tiff, generic));
        Assert.Equal(unit, routed.ResolutionUnits); Assert.Equal(x, routed.HorizontalResolution);
        generic.DpiX = 144; generic.DpiY = 120;
        OfficeImageMetadata overridden = OfficeImageMetadata.Read(OfficeRasterImageEncoder.Encode(pages[0], OfficeImageExportFormat.Tiff, generic));
        Assert.Equal(OfficeImageResolutionUnit.PixelsPerInch, overridden.ResolutionUnits); Assert.Equal(144D, overridden.HorizontalResolution); Assert.Equal(120D, overridden.VerticalResolution);
    }

    [Fact]
    public void MetadataRewriteRejectsBackingGrowthAndFinalCopyBeforeAllocatingBeyondBudget() {
        using var stream = new OfficeMetadataRewriteStream(OfficeRasterGuards.MaximumDecodedBytes - 512L, 256, System.Threading.CancellationToken.None);
        Assert.Throws<ArgumentException>(() => stream.Write(new byte[257], 0, 257));
        Assert.Equal(0L, stream.Length);
        using var finalCopy = new OfficeMetadataRewriteStream(OfficeRasterGuards.MaximumDecodedBytes - 500L, 256, System.Threading.CancellationToken.None);
        finalCopy.Write(new byte[256], 0, 256);
        Assert.Throws<ArgumentException>(() => finalCopy.ToArray());
        using var cancellation = new System.Threading.CancellationTokenSource();
        using var cancellable = new OfficeMetadataRewriteStream(0, 256, cancellation.Token);
        cancellable.WriteByte(42); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => cancellable.ToArray());
    }

    [Fact]
    public void MetadataReadOwnsTiffFieldsAfterCallerChangesTheWholeContainer() {
        byte[] tiff = OfficeTiffCodec.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        var source = OfficeImageMetadata.Read(tiff); source.SetExifValue(OfficeExifTag.Artist, "Independent"); source.XmpProfile = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>");
        byte[] annotated = OfficeImageMetadata.Apply(tiff, source);
        OfficeImageMetadata owned = OfficeImageMetadata.Read(annotated);
        byte[] expectedXmp = owned.XmpProfile!;
        Array.Clear(annotated, 0, annotated.Length);
        Assert.Equal("Independent", owned.GetExifValue(OfficeExifTag.Artist)!.Value);
        Assert.Equal(expectedXmp, owned.XmpProfile);
        Assert.Equal("Independent", OfficeImageMetadata.ParseExifProfile(owned.EncodeExifProfile()!).GetExifValue(OfficeExifTag.Artist)!.Value);
    }
    [Fact]
    public void GifProfilesPreserveOriginalImageAndTimingBlocksDuringSelectiveRemoval() {
        byte[] gif = Convert.FromBase64String("R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=");
        var metadata = OfficeImageMetadata.Read(gif);
        byte[] xmp = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>");
        metadata.XmpProfile = xmp;
        metadata.IccProfile = MinimalIcc();
        byte[] annotated = OfficeImageMetadata.Apply(gif, metadata);
        OfficeImageMetadata read = OfficeImageMetadata.Read(annotated);
        Assert.Equal(xmp, read.XmpProfile);
        Assert.Equal(metadata.IccProfile, read.IccProfile);
        OfficeImageMetadataRemovalResult selective = OfficeImageMetadata.Remove(annotated, OfficeImageMetadataProfileKinds.Xmp);
        Assert.Equal(OfficeImageMetadataProfileKinds.Xmp | OfficeImageMetadataProfileKinds.Icc, selective.PresentProfiles);
        Assert.Equal(OfficeImageMetadataProfileKinds.Xmp, selective.RemovedProfiles);
        Assert.Null(OfficeImageMetadata.Read(selective.EncodedBytes).XmpProfile);
        Assert.Equal(metadata.IccProfile, OfficeImageMetadata.Read(selective.EncodedBytes).IccProfile);
        Assert.Equal(gif, OfficeImageMetadata.Remove(annotated, OfficeImageMetadataProfileKinds.All).EncodedBytes);
    }

    [Fact]
    public void BitmapV5ProfileRemovalErasesProfileAndPreservesPixelPayload() {
        byte[] bmp = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.FromRgba(10, 30, 80, 90)), OfficeImageExportFormat.Bmp);
        var metadata = OfficeImageMetadata.Read(bmp);
        metadata.IccProfile = MinimalIcc();
        metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerMeter;
        metadata.HorizontalResolution = 6000;
        metadata.VerticalResolution = 5000;
        byte[] annotated = OfficeImageMetadata.Apply(bmp, metadata);
        Assert.Equal(metadata.IccProfile, OfficeImageMetadata.Read(annotated).IccProfile);
        Assert.Equal(OfficeImageResolutionUnit.PixelsPerMeter, OfficeImageMetadata.Read(annotated).ResolutionUnits);
        Assert.Equal(6000D, OfficeImageMetadata.Read(annotated).HorizontalResolution);
        Assert.True(OfficeRasterImageDecoder.TryDecode(annotated, out OfficeRasterImage? image));
        Assert.Equal(90, image!.GetPixel(1, 1).A);
        OfficeImageMetadataRemovalResult removal = OfficeImageMetadata.Remove(annotated, OfficeImageMetadataProfileKinds.Icc);
        Assert.Equal(OfficeImageMetadataProfileKinds.Icc, removal.RemovedProfiles);
        Assert.Null(OfficeImageMetadata.Read(removal.EncodedBytes).IccProfile);
        Assert.DoesNotContain("acsp", Encoding.ASCII.GetString(removal.EncodedBytes));
        int originalPixels = BitConverter.ToInt32(bmp, 10);
        int annotatedPixels = BitConverter.ToInt32(annotated, 10);
        Assert.Equal(SliceBytes(bmp, originalPixels, 24), SliceBytes(annotated, annotatedPixels, 24));
        Assert.Equal(SliceBytes(annotated, annotatedPixels, 24), SliceBytes(removal.EncodedBytes, annotatedPixels, 24));
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageResolutionUnit.AspectRatio, 2D, 3D)]
    [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageResolutionUnit.PixelsPerCentimeter, 60D, 50D)]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageResolutionUnit.AspectRatio, 2D, 3D)]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageResolutionUnit.PixelsPerMeter, 6000D, 5000D)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageResolutionUnit.PixelsPerCentimeter, 60D, 50D)]
    [InlineData(OfficeImageExportFormat.Tiff, OfficeImageResolutionUnit.PixelsPerCentimeter, 60D, 50D)]
    public void NativeResolutionUnitsSurviveMetadataRoundTrip(OfficeImageExportFormat format, OfficeImageResolutionUnit unit, double x, double y) {
        byte[] encoded = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue), format);
        var metadata = OfficeImageMetadata.Read(encoded);
        metadata.ResolutionUnits = unit;
        metadata.HorizontalResolution = x;
        metadata.VerticalResolution = y;
        OfficeImageMetadata actual = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(encoded, metadata));
        Assert.Equal(unit, actual.ResolutionUnits);
        Assert.Equal(x, actual.HorizontalResolution);
        Assert.Equal(y, actual.VerticalResolution);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Pbm)]
    [InlineData(OfficeImageExportFormat.Tga)]
    public void ProfileRemovalFromFormatsWithoutDefinedProfileCarriersIsIdentity(OfficeImageExportFormat format) {
        byte[] encoded = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.White), format);
        OfficeImageMetadataRemovalResult result = OfficeImageMetadata.Remove(encoded, OfficeImageMetadataProfileKinds.All);
        Assert.Equal(encoded, result.EncodedBytes);
        Assert.Equal(OfficeImageMetadataProfileKinds.None, result.PresentProfiles);
        Assert.Equal(OfficeImageMetadataProfileKinds.None, result.RemovedProfiles);
    }

    private static byte[] MinimalIcc() {
        byte[] bytes = new byte[132];
        bytes[3] = 132;
        Encoding.ASCII.GetBytes("acsp", 0, 4, bytes, 36);
        return bytes;
    }

    [Fact]
    public void PhysicalDpiConvertsOwnedResolutionUnitsAndLeavesAspectRatioUnitless() {
        var metadata = new OfficeImageMetadata { ResolutionUnits = OfficeImageResolutionUnit.PixelsPerMeter, HorizontalResolution = 6000, VerticalResolution = 5000 };
        Assert.Equal(152.4D, metadata.PhysicalDpiX!.Value, 8);
        Assert.Equal(127D, metadata.PhysicalDpiY!.Value, 8);
        metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerCentimeter;
        metadata.HorizontalResolution = 60;
        Assert.Equal(152.4D, metadata.PhysicalDpiX!.Value, 8);
        metadata.ResolutionUnits = OfficeImageResolutionUnit.AspectRatio;
        Assert.Null(metadata.PhysicalDpiX);
        Assert.Null(metadata.PhysicalDpiY);
    }

    [Fact]
    public void TranscodingProjectsUnsupportedProfilesAndReportsOmissionsWithoutMutatingSource() {
        var metadata = new OfficeImageMetadata { IptcProfile = new byte[] { 28, 2, 5, 0, 1, 65 }, XmpProfile = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>") };
        metadata.SetExifValue(OfficeExifTag.Software, "Camera");
        OfficeImageMetadata pngMetadata = metadata.PrepareForEncoding(OfficeImageFormat.Png, out OfficeImageMetadataProfileKinds omitted);
        Assert.Equal(OfficeImageMetadataProfileKinds.Iptc, omitted);
        Assert.Null(pngMetadata.IptcProfile);
        Assert.NotNull(metadata.IptcProfile);
        Assert.Equal(metadata.XmpProfile, pngMetadata.XmpProfile);
        byte[] png = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue), OfficeImageExportFormat.Png);
        Assert.Throws<NotSupportedException>(() => OfficeImageMetadata.Apply(png, metadata));
        OfficeImageMetadata actual = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(png, pngMetadata));
        Assert.Equal("Camera", actual.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.Equal(metadata.XmpProfile, actual.XmpProfile);
        pngMetadata.SetExifValue(OfficeExifTag.Software, "Edited");
        Assert.Equal("Camera", metadata.GetExifValue(OfficeExifTag.Software)!.Value);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Webp)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    public void ProfileEditsAndSelectiveRemovalPreserveDecodedPixels(OfficeImageExportFormat format) {
        var image = new OfficeRasterImage(3, 2, OfficeColor.FromRgb(50, 120, 220));
        image.SetPixel(1, 1, OfficeColor.FromRgb(220, 10, 60));
        byte[] original = OfficeRasterImageEncoder.Encode(image, format);
        Assert.True(OfficeRasterImageDecoder.TryDecode(original, out OfficeRasterImage? before));
        var metadata = OfficeImageMetadata.Read(original);
        metadata.SetExifValue(OfficeExifTag.Software, "Metadata test");
        metadata.SetExifValue(OfficeExifTag.ExposureTime, new OfficeRational(1, 125));
        metadata.SetExifValue(OfficeExifTag.GPSLatitude, new[] { new OfficeRational(51, 1), new OfficeRational(30, 1), new OfficeRational(0, 1) });
        metadata.XmpProfile = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>");
        byte[] encoded = OfficeImageMetadata.Apply(original, metadata);
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? after));
        Assert.Equal(before!.GetPixels(), after!.GetPixels());
        OfficeImageMetadata read = OfficeImageMetadata.Read(encoded);
        Assert.Equal("Metadata test", read.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.Equal(125U, Assert.IsType<OfficeRational>(read.GetExifValue(OfficeExifTag.ExposureTime)!.Value).Denominator);
        Assert.Equal(3, Assert.IsType<OfficeRational[]>(read.GetExifValue(OfficeExifTag.GPSLatitude)!.Value).Length);
        Assert.Equal(metadata.XmpProfile, read.XmpProfile);
        OfficeImageMetadataRemovalResult removal = OfficeImageMetadata.Remove(encoded, OfficeImageMetadataProfileKinds.Exif);
        Assert.Equal(OfficeImageMetadataProfileKinds.Exif, removal.RemovedProfiles);
        if (format == OfficeImageExportFormat.Tiff) {
            OfficeImageMetadata tiffRemaining = OfficeImageMetadata.Read(removal.EncodedBytes);
            Assert.Equal((ushort)1, tiffRemaining.GetExifValue(OfficeExifTag.Orientation)!.Value);
            Assert.Single(tiffRemaining.ExifValues);
        } else Assert.Null(OfficeImageMetadata.Read(removal.EncodedBytes).ExifProfile);
        Assert.Equal(metadata.XmpProfile, OfficeImageMetadata.Read(removal.EncodedBytes).XmpProfile);
        Assert.True(OfficeRasterImageDecoder.TryDecode(removal.EncodedBytes, out OfficeRasterImage? stripped));
        Assert.Equal(before.GetPixels(), stripped!.GetPixels());
        // The compressed image bytes themselves, rather than only their decoded values, are preserved.
        if (format == OfficeImageExportFormat.Png) Assert.Equal(PngChunks(original, "IDAT"), PngChunks(encoded, "IDAT"));
        if (format == OfficeImageExportFormat.Jpeg) Assert.Equal(JpegScan(original), JpegScan(encoded));
        if (format == OfficeImageExportFormat.Webp) Assert.Equal(WebpChunks(original, "VP8L"), WebpChunks(encoded, "VP8L"));
    }

    [Fact]
    public void ExifTypedValuesPreserveByteOrderAndOpaqueOffsetsDuringEdits() {
        byte[] profile = new byte[64];
        profile[0] = 77; profile[1] = 77; Put(profile, 2, 42, 2, false); Put(profile, 4, 8, 4, false); Put(profile, 8, 2, 2, false);
        Put(profile, 10, 305, 2, false); Put(profile, 12, 2, 2, false); Put(profile, 14, 5, 4, false); Put(profile, 18, 40, 4, false);
        Put(profile, 22, 65000, 2, false); Put(profile, 24, 7, 2, false); Put(profile, 26, 8, 4, false); Put(profile, 30, 48, 4, false);
        Buffer.BlockCopy(Encoding.ASCII.GetBytes("old!\0"), 0, profile, 40, 5); for (int i = 0; i < 8; i++) profile[48 + i] = (byte)(100 + i);
        OfficeImageMetadata metadata = OfficeImageMetadata.ParseExifProfile(profile);
        metadata.SetExifValue(OfficeExifTag.Software, "replacement"); metadata.SetExifValue(OfficeExifTag.Artist, "Creator");
        byte[] edited = metadata.EncodeExifProfile()!;
        Assert.Equal(SliceBytes(profile, 48, 8), SliceBytes(edited, 48, 8));
        Assert.DoesNotContain("old!", Encoding.ASCII.GetString(edited));
        OfficeImageMetadata read = OfficeImageMetadata.ParseExifProfile(edited);
        Assert.Equal("replacement", read.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.Equal("Creator", read.GetExifValue(OfficeExifTag.Artist)!.Value);
        Assert.True(read.RemoveExifValue(OfficeExifTag.Software));
        Assert.DoesNotContain("replacement", Encoding.ASCII.GetString(read.EncodeExifProfile()!));
    }

    [Fact]
    public void AliasedExifValuesAreRejectedRatherThanErasingAnotherField() {
        byte[] profile = new byte[48]; profile[0] = 73; profile[1] = 73; Put(profile, 2, 42, 2, true); Put(profile, 4, 8, 4, true); Put(profile, 8, 2, 2, true);
        for (int i = 0; i < 2; i++) { int entry = 10 + i * 12; Put(profile, entry, i == 0 ? 305U : 315U, 2, true); Put(profile, entry + 2, 2, 2, true); Put(profile, entry + 4, 5, 4, true); Put(profile, entry + 8, 40, 4, true); }
        Buffer.BlockCopy(Encoding.ASCII.GetBytes("text\0"), 0, profile, 40, 5);
        OfficeImageMetadata metadata = OfficeImageMetadata.ParseExifProfile(profile); metadata.RemoveExifValue(OfficeExifTag.Software);
        Assert.Throws<FormatException>(() => metadata.EncodeExifProfile());
        Assert.Equal("text", metadata.GetExifValue(OfficeExifTag.Artist)!.Value);
    }

    [Fact]
    public void ProfilesAndArrayValuesOwnTheirStorage() {
        var metadata = new OfficeImageMetadata(); byte[] xmp = new byte[] { 1, 2, 3 }; metadata.XmpProfile = xmp; xmp[0] = 9;
        byte[] copied = metadata.XmpProfile!; copied[1] = 9; Assert.Equal(new byte[] { 1, 2, 3 }, metadata.XmpProfile);
        ushort[] sensitivity = { 100, 200 }; metadata.SetExifValue(OfficeExifTag.ISOSpeedRatings, sensitivity); sensitivity[0] = 999;
        Assert.Equal(new ushort[] { 100, 200 }, Assert.IsType<ushort[]>(metadata.GetExifValue(OfficeExifTag.ISOSpeedRatings)!.Value));
    }

    [Fact]
    public void RemovingAnInlineValueAfterAnotherEditErasesSupersededDirectoryCopies() {
        var metadata = new OfficeImageMetadata(); metadata.SetExifValue(OfficeExifTag.Software, "Old");
        metadata = OfficeImageMetadata.ParseExifProfile(metadata.EncodeExifProfile()!);
        metadata.SetExifValue(OfficeExifTag.Artist, "Creator");
        metadata = OfficeImageMetadata.ParseExifProfile(metadata.EncodeExifProfile()!);
        Assert.True(metadata.RemoveExifValue(OfficeExifTag.Software));
        Assert.DoesNotContain("Old", Encoding.ASCII.GetString(metadata.EncodeExifProfile()!));
    }

    [Fact]
    public void TiffMetadataRemovalPreservesAllPagesAndErasesOpaquePayloads() {
        byte[] tiff = OfficeTiffCodec.EncodePages(new[] { new OfficeRasterImage(3, 2, OfficeColor.Red), new OfficeRasterImage(2, 3, OfficeColor.Blue) });
        var metadata = OfficeImageMetadata.Read(tiff);
        var maker = new OfficeExifTag(37500, OfficeExifDataType.Undefined, OfficeExifDirectory.Exif);
        metadata.SetExifValue(maker, Encoding.ASCII.GetBytes("private maker payload")); metadata.SetExifValue(OfficeExifTag.Artist, "Private creator");
        metadata.XmpProfile = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>");
        byte[] annotated = OfficeImageMetadata.Apply(tiff, metadata);
        OfficeImageMetadata read = OfficeImageMetadata.Read(annotated);
        Assert.True(read.RequiresOriginalTiffContainer); Assert.Throws<NotSupportedException>(() => read.EncodeExifProfile());
        Assert.Throws<NotSupportedException>(() => read.PrepareForEncoding(OfficeImageFormat.Png, out _));
        Assert.Throws<NotSupportedException>(() => read.PrepareForEncoding(OfficeImageFormat.Tiff, out _));
        read.SetExifValue(OfficeExifTag.Software, "Safe edit"); byte[] edited = OfficeImageMetadata.Apply(annotated, read);
        Assert.Contains("private maker payload", Encoding.ASCII.GetString(edited));
        Assert.Equal(annotated, OfficeImageMetadata.Remove(annotated, OfficeImageMetadataProfileKinds.None).EncodedBytes);
        byte[] stripped = OfficeImageMetadata.Remove(edited, OfficeImageMetadataProfileKinds.Exif).EncodedBytes;
        Assert.DoesNotContain("private maker payload", Encoding.ASCII.GetString(stripped)); Assert.DoesNotContain("Private creator", Encoding.ASCII.GetString(stripped));
        OfficeImageMetadata remaining = OfficeImageMetadata.Read(stripped);
        Assert.Single(remaining.ExifValues); Assert.Equal((ushort)1, remaining.GetExifValue(OfficeExifTag.Orientation)!.Value);
        Assert.Equal(metadata.XmpProfile, remaining.XmpProfile);
        Assert.True(OfficeRasterContainerInspector.TryInspect(stripped, out OfficeRasterContainerInfo? pages)); Assert.Equal(2, pages!.Count);
        for (int index = 0; index < 2; index++) { var options = new OfficeRasterDecodeOptions { FrameIndex = index }; Assert.True(OfficeRasterImageDecoder.TryDecode(annotated, options, out OfficeRasterImage? before, out _)); Assert.True(OfficeRasterImageDecoder.TryDecode(stripped, options, out OfficeRasterImage? after, out _)); Assert.Equal(before!.GetPixels(), after!.GetPixels()); }
    }
    private static void Put(byte[] bytes, int at, uint value, int size, bool little) { for (int i = 0; i < size; i++) bytes[at + i] = (byte)(value >> (8 * (little ? i : size - i - 1))); }
    private static byte[] PngChunks(byte[] bytes, string wanted) { using var result = new MemoryStream(); for (int cursor = 8; cursor < bytes.Length;) { int length = (int)Read(bytes, cursor, false); if (Encoding.ASCII.GetString(bytes, cursor + 4, 4) == wanted) result.Write(bytes, cursor + 8, length); cursor += 12 + length; } return result.ToArray(); }
    private static byte[] WebpChunks(byte[] bytes, string wanted) { using var result = new MemoryStream(); for (int cursor = 12; cursor < bytes.Length;) { int length = (int)Read(bytes, cursor + 4, true); if (Encoding.ASCII.GetString(bytes, cursor, 4) == wanted) result.Write(bytes, cursor + 8, length); cursor += 8 + length + (length & 1); } return result.ToArray(); }
    private static byte[] JpegScan(byte[] bytes) { int cursor = 2; while (cursor < bytes.Length) { int marker = bytes[cursor + 1]; if (marker == 0xDA) return SliceBytes(bytes, cursor, bytes.Length - cursor); int size = bytes[cursor + 2] << 8 | bytes[cursor + 3]; cursor += size + 2; } throw new InvalidDataException(); }
    private static byte[] SliceBytes(byte[] bytes, int at, int count) { var result = new byte[count]; Buffer.BlockCopy(bytes, at, result, 0, count); return result; }
    private static uint Read(byte[] bytes, int at, bool little) => little ? (uint)(bytes[at] | bytes[at + 1] << 8 | bytes[at + 2] << 16 | bytes[at + 3] << 24) : (uint)(bytes[at] << 24 | bytes[at + 1] << 16 | bytes[at + 2] << 8 | bytes[at + 3]);
}
