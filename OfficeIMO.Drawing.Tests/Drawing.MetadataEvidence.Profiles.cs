using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed partial class ImageMetadataEvidenceTests {
    [Fact]
    public void XmpPreservesNamespacesCollectionOrderLanguagesAndResourceUris() {
        byte[] packet = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/' xmlns:r='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:d='http://purl.org/dc/elements/1.1/' xmlns:p='http://ns.adobe.com/xap/1.0/'>" +
            "<r:RDF><r:Description p:CreatorTool='Camera'><d:creator><r:Seq><r:li>Alice</r:li><r:li>Bob</r:li></r:Seq></d:creator>" +
            "<d:title><r:Alt><r:li xml:lang='x-default'>Photo</r:li><r:li xml:lang='pl'>Zdjęcie</r:li></r:Alt></d:title>" +
            "<d:source r:resource='https://example.test/photo'/><d:relation r:parseType='Resource'><d:title>Nested</d:title></d:relation></r:Description></r:RDF></x:xmpmeta>");
        OfficeXmpInspection result = OfficeXmpInspector.Inspect(packet);
        Assert.Equal(6, result.Properties.Count);
        Assert.Equal(new[] { "Alice", "Bob" }, result.Properties.Where(property => property.Name == "creator").Select(property => property.Value));
        Assert.Equal(new int?[] { 0, 1 }, result.Properties.Where(property => property.Name == "creator").Select(property => property.ItemIndex));
        Assert.Contains(result.Properties, property => property.Kind == OfficeXmpValueKind.AlternativeItem && property.Language == "pl" && property.Value == "Zdjęcie");
        Assert.Contains(result.Properties, property => property.Kind == OfficeXmpValueKind.Resource && property.Value == "https://example.test/photo");
        Assert.Contains("xmp:structured-property-omitted", result.Issues);
        Assert.DoesNotContain(result.Properties, property => property.Value == "Nested");
    }

    [Fact]
    public void XmpRejectsDtdAndReportsRetentionAndNodeLimits() {
        OfficeXmpInspection hostile = OfficeXmpInspector.Inspect(Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY secret SYSTEM 'file:///unreadable'>]><x>&secret;</x>"));
        Assert.Contains("xmp:invalid-xml", hostile.Issues); Assert.Empty(hostile.Properties);
        byte[] packet = Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/'><rdf:Description><dc:title>Long title</dc:title><dc:creator>Alice</dc:creator></rdf:Description></rdf:RDF>");
        OfficeXmpInspection limited = OfficeXmpInspector.Inspect(packet, new OfficeImageMetadataInspectionOptions { MaximumXmpProperties = 1, MaximumValueCharacters = 4 });
        Assert.Equal("Long", Assert.Single(limited.Properties).Value);
        Assert.True(limited.Properties[0].IsTruncated);
        Assert.Contains("xmp:property-limit", limited.Issues);
        Assert.Contains("xmp:xml-node-limit", OfficeXmpInspector.Inspect(packet, new OfficeImageMetadataInspectionOptions { MaximumXmlNodes = 1 }).Issues);
    }

    [Fact]
    public void IccValidEnvelopeAndUnsupportedTransformAreDistinctFromCorruption() {
        byte[] profile = new byte[132]; Put(profile, 0, 132, 4); profile[8] = 4; profile[9] = 0x20;
        Encoding.ASCII.GetBytes("mntrRGB XYZ ").CopyTo(profile, 12); Encoding.ASCII.GetBytes("acsp").CopyTo(profile, 36);
        OfficeIccProfileInspection result = OfficeIccProfileInspector.Inspect(profile);
        Assert.True(result.IsStructurallyValid); Assert.False(result.HasSupportedInputTransform);
        Assert.Equal("RGB ", result.DeviceColorSpaceSignature); Assert.Equal(4, result.MajorVersion);
        Assert.Contains("icc:input-transform-unsupported", result.Issues);
        profile[36] = 0;
        OfficeIccProfileInspection invalid = OfficeIccProfileInspector.Inspect(profile);
        Assert.False(invalid.IsStructurallyValid); Assert.Null(invalid.HasSupportedInputTransform);
        Assert.Contains("icc:invalid-structure", invalid.Issues);
    }

    [Fact]
    public void IccTagRetentionDoesNotSkipValidationOfUnretainedTagRanges() {
        byte[] profile = new byte[172]; Put(profile, 0, 172, 4); Encoding.ASCII.GetBytes("acsp").CopyTo(profile, 36);
        Put(profile, 128, 2, 4);
        Encoding.ASCII.GetBytes("desc").CopyTo(profile, 132); Put(profile, 136, 156, 4); Put(profile, 140, 8, 4); Encoding.ASCII.GetBytes("text").CopyTo(profile, 156);
        Encoding.ASCII.GetBytes("cprt").CopyTo(profile, 144); Put(profile, 148, 164, 4); Put(profile, 152, 8, 4); Encoding.ASCII.GetBytes("text").CopyTo(profile, 164);
        var options = new OfficeImageMetadataInspectionOptions { MaximumIccTags = 1 };
        OfficeIccProfileInspection result = OfficeIccProfileInspector.Inspect(profile, options);
        Assert.True(result.IsStructurallyValid); Assert.Equal(2U, result.DeclaredTagCount); Assert.Single(result.Tags);
        Assert.Contains("icc:tag-retention-limit", result.Issues);
        Put(profile, 148, 10000, 4);
        Assert.False(OfficeIccProfileInspector.Inspect(profile, options).IsStructurallyValid);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void JpegExtendedXmpUsesDigestBoundAssembly(bool tampered) {
        byte[] extended = Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/'><rdf:Description dc:creator='Independent packet'/></rdf:RDF>");
        string guid;
        using (var md5 = MD5.Create()) guid = BitConverter.ToString(md5.ComputeHash(extended)).Replace("-", string.Empty);
        byte[] standard = Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:n='http://ns.adobe.com/xmp/note/'><rdf:Description n:HasExtendedXMP='" + guid + "'/></rdf:RDF>");
        byte[] prefix = Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0" + guid);
        var extension = new byte[prefix.Length + 8 + extended.Length]; prefix.CopyTo(extension, 0); Put(extension, prefix.Length, (uint)extended.Length, 4); extended.CopyTo(extension, prefix.Length + 8);
        if (tampered) extension[extension.Length - 1] ^= 1;
        byte[] jpeg = Jpeg(JpegSegment(0xE1, Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0").Concat(standard).ToArray()), JpegSegment(0xE1, extension));
        OfficeImageMetadataInspection result = OfficeImageMetadata.Inspect(jpeg);
        if (tampered) { Assert.Contains("xmp:extended-digest-mismatch", result.Issues); Assert.DoesNotContain(result.XmpProperties, property => property.Name == "creator"); }
        else Assert.Contains(result.XmpProperties, property => property.Name == "creator" && property.Value == "Independent packet" && property.Source.Contains("ExtendedXMP"));
    }

    [Fact]
    public void AuthoredResolutionNeverConfusesAbsentDensityWithNinetySixDpiDefault() {
        byte[] absent = Jpeg();
        Assert.Equal(96, OfficeImageReader.Identify(absent).DpiX);
        OfficeImageMetadataInspection missing = OfficeImageMetadata.Inspect(absent);
        Assert.False(missing.Resolution.IsDeclared); Assert.Null(missing.Resolution.Resolution);
        byte[] jfif = { 74, 70, 73, 70, 0, 1, 2, 1, 1, 44, 1, 44, 0, 0 };
        OfficeImageMetadataInspection authored = OfficeImageMetadata.Inspect(Jpeg(JpegSegment(0xE0, jfif)));
        Assert.True(authored.Resolution.IsDeclared); Assert.True(authored.Resolution.IsValid);
        Assert.Equal(300, authored.Resolution.Resolution!.PhysicalDpiX);
        byte[] contradictory = Exif(true, Field(OfficeExifDirectory.Image, 282, 5, Rationals(true, (96, 1))), Field(OfficeExifDirectory.Image, 283, 5, Rationals(true, (96, 1))));
        OfficeImageMetadataInspection conflict = OfficeImageMetadata.Inspect(Jpeg(JpegSegment(0xE0, jfif), JpegSegment(0xE1, AsciiPrefixExif(contradictory))));
        Assert.Contains("resolution:conflicting-carriers", conflict.Issues); Assert.Null(conflict.Resolution.Resolution);
    }

    [Fact]
    public void SeparateExifCarriersCannotContributePiecesOfOneTimestampOrCoordinate() {
        byte[] capture = Exif(true, Field(OfficeExifDirectory.Exif, 36867, 2, Ascii("2026:10:13 12:34:56")),
            Field(OfficeExifDirectory.Gps, 1, 2, Ascii("N")), Field(OfficeExifDirectory.Gps, 2, 5, Rationals(true, (51, 1), (30, 1), (0, 1))));
        byte[] other = Exif(true, Field(OfficeExifDirectory.Exif, 36881, 2, Ascii("+02:00")),
            Field(OfficeExifDirectory.Gps, 3, 2, Ascii("E")), Field(OfficeExifDirectory.Gps, 4, 5, Rationals(true, (1, 1), (15, 1), (0, 1))));
        OfficeImageMetadataInspection result = OfficeImageMetadata.Inspect(Jpeg(JpegSegment(0xE1, AsciiPrefixExif(capture)), JpegSegment(0xE1, AsciiPrefixExif(other))));
        Assert.Contains("exif:duplicate-carriers", result.Issues);
        Assert.False(Assert.Single(result.Timestamps).HasKnownTimezone);
        Assert.Null(result.GpsCoordinate);
    }

    [Fact]
    public void OmittedDensityFieldsDoNotEstablishAuthoredDensityAbsence() {
        byte[] bytes = Exif(true, Field(OfficeExifDirectory.Image, 271, 2, Ascii("Camera")),
            Field(OfficeExifDirectory.Image, 282, 5, Rationals(true, (300, 1))), Field(OfficeExifDirectory.Image, 283, 5, Rationals(true, (300, 1))));
        OfficeImageMetadataInspection result = OfficeImageMetadata.InspectExifProfile(bytes, new OfficeImageMetadataInspectionOptions { MaximumExifFields = 1 });
        Assert.Contains("exif:field-retention-limit", result.Issues);
        Assert.False(result.Resolution.IsPresenceKnown); Assert.Null(result.Resolution.Resolution);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Webp)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    public void ExistingImageCarriersFeedTheSameXmpInspectionContract(OfficeImageExportFormat format) {
        byte[] encoded = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue), format);
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(encoded);
        metadata.XmpProfile = Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/'><rdf:Description dc:creator='Camera'/></rdf:RDF>");
        byte[] annotated = OfficeImageMetadata.Apply(encoded, metadata);
        OfficeImageMetadataInspection result = OfficeImageMetadata.Inspect(annotated);
        Assert.Contains(result.XmpProperties, property => property.Name == "creator" && property.Value == "Camera");
    }

    private static byte[] AsciiPrefixExif(byte[] bytes) => new byte[] { 69, 120, 105, 102, 0, 0 }.Concat(bytes).ToArray();
    private static byte[] Jpeg(params byte[][] segments) => new byte[] { 255, 216 }.Concat(segments.SelectMany(segment => segment)).Concat(new byte[] {
        255, 192, 0, 17, 8, 0, 1, 0, 1, 3, 1, 17, 0, 2, 17, 0, 3, 17, 0, 255, 217 }).ToArray();
    private static byte[] JpegSegment(byte marker, byte[] payload) {
        var segment = new byte[payload.Length + 4]; segment[0] = 255; segment[1] = marker; Put(segment, 2, (uint)(payload.Length + 2), 2); payload.CopyTo(segment, 4); return segment;
    }
}
