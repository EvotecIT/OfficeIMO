using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed partial class ImageMetadataEvidenceTests {
    [Fact]
    public void OmittedStructuralValuesDoNotInvalidateAnOtherwiseValidSharedRangeProfile() {
        // Every entry aliases one bounded source range. Inspection must retain descriptors,
        // rather than decode the omitted range repeatedly and hit the ordinary-read budget.
        const int entries = 520, valueBytes = 256 * 1024;
        int payload = 14 + entries * 12;
        byte[] bytes = new byte[payload + valueBytes];
        bytes[0] = bytes[1] = 73; Put(bytes, 2, 42, 2, true); Put(bytes, 4, 8, 4, true); Put(bytes, 8, entries, 2, true);
        ushort[] ids = { 513, 514, 34665, 34853, 40965, 330, 34665 };
        ushort[] types = { 4, 4, 4, 4, 4, 7, 2 };
        for (int i = 0; i < entries; i++) {
            int entry = 10 + i * 12;
            ushort type = types[i % types.Length];
            Put(bytes, entry, ids[i % ids.Length], 2, true); Put(bytes, entry + 2, type, 2, true);
            Put(bytes, entry + 4, (uint)(type == 4 ? valueBytes / 4 : valueBytes), 4, true); Put(bytes, entry + 8, (uint)payload, 4, true);
        }
        OfficeImageMetadataInspection result = OfficeImageMetadata.InspectExifProfile(bytes,
            new OfficeImageMetadataInspectionOptions { MaximumExifFields = 7, MaximumExifValueBytes = 4, MaximumRetainedExifBytes = 4 });
        Assert.DoesNotContain("exif:invalid-profile", result.Issues);
        Assert.Contains("exif:field-retention-limit", result.Issues);
        Assert.Equal(7, result.ExifFields.Count);
        Assert.All(result.ExifFields, field => { Assert.Null(field.Value); Assert.Equal((uint)(field.RawDataType == 4 ? valueBytes / 4 : valueBytes), field.Count); Assert.Equal(payload, field.ValueOffset); });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DuplicateGifXmpCarriersAreLocatedAndNeverReportedComplete(bool rawXmp) {
        byte[] first = ReviewXmp("First"), second = ReviewXmp("Second");
        byte[] gif = ReviewGif(ReviewGifApplication("XMP DataXMP", first, rawXmp), ReviewGifApplication("XMP DataXMP", second, rawXmp));
        OfficeImageMetadataInspection result = OfficeImageMetadata.Inspect(gif);
        Assert.Contains("xmp:duplicate-carriers", result.Issues); Assert.False(result.IsComplete);
        Assert.Equal(new[] { "First", "Second" }, result.XmpProperties.Select(property => property.Value));
        Assert.Equal(2, result.XmpProperties.Select(property => property.Source).Distinct().Count());
        Assert.Equal(first, OfficeImageMetadata.Read(gif).XmpProfile);
        Assert.True(OfficeImageMetadata.Inspect(ReviewGif(ReviewGifApplication("XMP DataXMP", first, rawXmp))).IsComplete);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GifProfileOmissionsStillCountOccurrencesAndInspectLaterCarriers(bool rawXmp) {
        byte[] smallIcc = new byte[132]; Put(smallIcc, 0, 132, 4); Encoding.ASCII.GetBytes("acsp").CopyTo(smallIcc, 36);
        byte[] largeIcc = new byte[512]; Put(largeIcc, 0, 512, 4); Encoding.ASCII.GetBytes("acsp").CopyTo(largeIcc, 36);
        byte[] gif = ReviewGif(ReviewGifApplication("XMP DataXMP", ReviewXmp("First"), rawXmp),
            ReviewGifApplication("XMP DataXMP", ReviewXmp(new string('X', 512)), rawXmp),
            ReviewGifApplication("ICCRGBG1012", smallIcc), ReviewGifApplication("ICCRGBG1012", largeIcc),
            ReviewGifApplication("XMP DataXMP", ReviewXmp("Last"), rawXmp));
        OfficeImageMetadataInspection result = OfficeImageMetadata.Inspect(gif,
            new OfficeImageMetadataInspectionOptions { MaximumProfileBytes = 256, MaximumExifValueBytes = 256 });
        Assert.Contains("image:profile-byte-limit", result.Issues);
        Assert.Contains("xmp:duplicate-carriers", result.Issues); Assert.Contains("icc:duplicate-carriers", result.Issues);
        Assert.Equal(new[] { "First", "Last" }, result.XmpProperties.Select(property => property.Value));
        Assert.Null(result.IccProfile);
    }

    [Fact]
    public void DuplicateGifIccCarriersDoNotSelectOneEffectiveProfile() {
        byte[] first = new byte[132]; Put(first, 0, 132, 4); Encoding.ASCII.GetBytes("acsp").CopyTo(first, 36);
        byte[] second = (byte[])first.Clone(); second[4] = 1;
        byte[] gif = ReviewGif(ReviewGifApplication("ICCRGBG1012", first), ReviewGifApplication("ICCRGBG1012", second));
        OfficeImageMetadataInspection result = OfficeImageMetadata.Inspect(gif);
        Assert.Contains("icc:duplicate-carriers", result.Issues); Assert.Null(result.IccProfile);
        Assert.Equal(first, OfficeImageMetadata.Read(gif).IccProfile);
    }

    [Fact]
    public void DuplicateGpsDirectoryPointersCannotEstablishOneNormalizedCoordinate() {
        byte[] first = ReviewGpsExif(51), second = ReviewGpsExif(52);
        int firstGps = BitConverter.ToInt32(first, 18), secondGps = BitConverter.ToInt32(second, 18);
        int secondAt = first.Length, newRoot = secondAt + second.Length - secondGps;
        byte[] bytes = new byte[newRoot + 30]; first.CopyTo(bytes, 0);
        Buffer.BlockCopy(second, secondGps, bytes, secondAt, second.Length - secondGps);
        for (int i = 0; i < 4; i++) {
            int entry = secondAt + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, entry + 2) == 5) Put(bytes, entry + 8, (uint)(BitConverter.ToInt32(bytes, entry + 8) + secondAt - secondGps), 4, true);
        }
        Put(bytes, 4, (uint)newRoot, 4, true); Put(bytes, newRoot, 2, 2, true);
        for (int i = 0; i < 2; i++) {
            int entry = newRoot + 2 + i * 12;
            Put(bytes, entry, 34853, 2, true); Put(bytes, entry + 2, 4, 2, true); Put(bytes, entry + 4, 1, 4, true);
            Put(bytes, entry + 8, (uint)(i == 0 ? firstGps : secondAt), 4, true);
        }
        OfficeImageMetadataInspection result = OfficeImageMetadata.InspectExifProfile(bytes);
        Assert.DoesNotContain("exif:invalid-profile", result.Issues);
        Assert.Equal(2, result.ExifFields.Count(field => field.TagId == 34853));
        Assert.Contains("exif:directory-pointer-ambiguous", result.Issues); Assert.Null(result.GpsCoordinate);
        OfficeImageMetadataInspection limited = OfficeImageMetadata.InspectExifProfile(bytes, new OfficeImageMetadataInspectionOptions { MaximumExifFields = 5 });
        Assert.Contains("exif:field-retention-limit", limited.Issues); Assert.Null(limited.GpsCoordinate);
        Assert.NotNull(OfficeImageMetadata.InspectExifProfile(first).GpsCoordinate);
    }

    private static byte[] ReviewGpsExif(uint latitude) => Exif(true,
        Field(OfficeExifDirectory.Gps, 1, 2, Ascii("N")), Field(OfficeExifDirectory.Gps, 2, 5, Rationals(true, (latitude, 1), (30, 1), (0, 1))),
        Field(OfficeExifDirectory.Gps, 3, 2, Ascii("E")), Field(OfficeExifDirectory.Gps, 4, 5, Rationals(true, (1, 1), (15, 1), (0, 1))));
    private static byte[] ReviewXmp(string creator) => Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/'><rdf:Description dc:creator='" + creator + "'/></rdf:RDF>");
    private static byte[] ReviewGif(params byte[][] applications) {
        byte[] image = Convert.FromBase64String("R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=");
        return image.Take(19).Concat(applications.SelectMany(application => application)).Concat(image.Skip(19)).ToArray();
    }
    private static byte[] ReviewGifApplication(string identifier, byte[] profile, bool rawXmp = false) {
        using var output = new MemoryStream(); output.Write(new byte[] { 0x21, 0xFF, 11 }); output.Write(Encoding.ASCII.GetBytes(identifier));
        if (rawXmp) output.Write(profile);
        else for (int at = 0; at < profile.Length; at += 255) { int length = Math.Min(255, profile.Length - at); output.WriteByte((byte)length); output.Write(profile, at, length); }
        if (identifier == "XMP DataXMP") { output.WriteByte(1); for (int value = 255; value >= 0; value--) output.WriteByte((byte)value); }
        output.WriteByte(0); return output.ToArray();
    }
}
