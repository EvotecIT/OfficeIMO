using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed partial class ImageMetadataEvidenceTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void IndependentExifBytesRetainCountsDuplicatesOffsetsAndNormalizedCapture(bool little) {
        byte[] bytes = Exif(little,
            Field(OfficeExifDirectory.Image, 271, 2, Ascii("Camera")),
            Field(OfficeExifDirectory.Image, 271, 2, Ascii("Other")),
            Field(OfficeExifDirectory.Exif, 36867, 2, Ascii("2026:10:13 12:34:56")),
            Field(OfficeExifDirectory.Exif, 36881, 2, Ascii("+02:00")),
            Field(OfficeExifDirectory.Exif, 37521, 2, Ascii("125")),
            Field(OfficeExifDirectory.Gps, 1, 2, Ascii("N")),
            Field(OfficeExifDirectory.Gps, 2, 5, Rationals(little, (51, 1), (30, 1), (0, 1))),
            Field(OfficeExifDirectory.Gps, 3, 2, Ascii("W")),
            Field(OfficeExifDirectory.Gps, 4, 5, Rationals(little, (1, 1), (15, 1), (0, 1))),
            Field(OfficeExifDirectory.Exif, 37510, 7, Array.Empty<byte>()));
        byte[] original = (byte[])bytes.Clone();
        OfficeImageMetadataInspection result = OfficeImageMetadata.InspectExifProfile(bytes);
        Assert.Equal(2, result.ExifFields.Count(field => field.TagId == 271));
        Assert.Contains("exif:duplicate-tag", result.Issues);
        OfficeExifFieldEvidence empty = Assert.Single(result.ExifFields, field => field.TagId == 37510);
        Assert.Equal(0U, empty.Count);
        Assert.Empty(Assert.IsType<byte[]>(empty.Value!.Value));
        Assert.All(result.ExifFields, field => Assert.True(field.EntryOffset >= field.DirectoryOffset + 2));
        OfficeExifTimestamp timestamp = Assert.Single(result.Timestamps);
        Assert.Equal("2026:10:13 12:34:56", timestamp.RawValue);
        Assert.Equal(new DateTimeOffset(2026, 10, 13, 12, 34, 56, 125, TimeSpan.FromHours(2)), timestamp.Instant);
        Assert.Equal(DateTimeKind.Unspecified, timestamp.LocalDateTime!.Value.Kind);
        Assert.Equal(51.5, result.GpsCoordinate!.Latitude);
        Assert.Equal(-1.25, result.GpsCoordinate.Longitude);
        Assert.Equal(original, bytes);
    }

    [Theory]
    [InlineData(null, null, true, false)]
    [InlineData("+02:00", "125", true, true)]
    [InlineData("+14:01", null, false, false)]
    [InlineData("+00:00", "not-a-fraction", false, false)]
    [InlineData(null, "123456789", true, false)]
    [InlineData("   :  ", null, true, false)]
    public void CaptureTimestampDoesNotInventUtcOrAcceptMalformedCompanions(string? offset, string? subseconds, bool valid, bool timezone) {
        OfficeExifTimestamp value = OfficeExifTimestamp.Parse("2026:10:13 12:34:56", offset, subseconds);
        Assert.Equal(valid, value.IsValid);
        Assert.Equal(timezone, value.HasKnownTimezone);
        Assert.Equal(offset, value.RawOffset);
        Assert.Equal(subseconds, value.RawSubseconds);
        Assert.Equal(DateTimeKind.Unspecified, value.LocalDateTime!.Value.Kind);
        if (!timezone) Assert.Null(value.Instant);
        if (subseconds == "123456789") Assert.Equal("exif:subsecond-precision-truncated", value.Issue);
        Assert.False(OfficeExifTimestamp.Parse("2026:02:30 12:34:56").IsValid);
    }

    [Theory]
    [InlineData("1")]
    [InlineData("9999999")]
    [InlineData("999999999")]
    public void MaximumWholeSecondAcceptsRepresentableFractionalTicks(string subseconds) {
        OfficeExifTimestamp value = OfficeExifTimestamp.Parse("9999:12:31 23:59:59", "+00:00", subseconds);
        Assert.True(value.IsValid); Assert.True(value.HasKnownTimezone);
        Assert.Equal(DateTimeKind.Unspecified, value.LocalDateTime!.Value.Kind);
        Assert.InRange(value.LocalDateTime.Value.Ticks, DateTime.MaxValue.Ticks - 9999999, DateTime.MaxValue.Ticks);
        Assert.Equal(subseconds, value.RawSubseconds);
        if (subseconds.Length > 7) Assert.Equal("exif:subsecond-precision-truncated", value.Issue);
    }

    [Theory]
    [InlineData(51, 0, "N")]
    [InlineData(51, 1, "north")]
    [InlineData(91, 1, "N")]
    public void GpsRejectsInvalidRationalsReferencesAndOutOfRangeCoordinates(uint degrees, uint denominator, string reference) {
        Assert.False(OfficeExifGpsCoordinate.TryCreate(new[] { new OfficeRational(degrees, denominator), new OfficeRational(0, 1), new OfficeRational(0, 1) }, reference,
            new[] { new OfficeRational(1, 1), new OfficeRational(0, 1), new OfficeRational(0, 1) }, "E", out _));
    }

    [Fact]
    public void InspectionRetainsOmittedValueDescriptorAndRefusesAmbiguousNormalization() {
        byte[] profile = Exif(true,
            Field(OfficeExifDirectory.Exif, 36867, 2, Ascii("2026:10:13 12:34:56")),
            Field(OfficeExifDirectory.Exif, 36881, 2, Ascii("+02:00")),
            Field(OfficeExifDirectory.Exif, 36881, 2, Ascii("+03:00")),
            Field(OfficeExifDirectory.Exif, 37500, 7, Enumerable.Repeat((byte)42, 64).ToArray()));
        var options = new OfficeImageMetadataInspectionOptions { MaximumExifValueBytes = 32 };
        OfficeImageMetadataInspection result = OfficeImageMetadata.InspectExifProfile(profile, options);
        OfficeExifFieldEvidence opaque = Assert.Single(result.ExifFields, field => field.TagId == 37500);
        Assert.Equal(64U, opaque.Count); Assert.Null(opaque.Value);
        Assert.Equal("exif:value-byte-limit", opaque.Omission);
        Assert.Empty(result.Timestamps);
        Assert.Contains("exif:timestamp-companion-invalid-or-ambiguous", result.Issues);
        Assert.Contains("exif:makernote-semantics-not-inspected", result.Issues);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeImageMetadata.InspectExifProfile(profile, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void TruncatedExifIsDiagnosedAndAggregateRetentionRemainsBounded() {
        byte[] profile = Exif(true, Field(OfficeExifDirectory.Image, 271, 2, Ascii("Camera")), Field(OfficeExifDirectory.Image, 272, 2, Ascii("Model")));
        OfficeImageMetadataInspection truncated = OfficeImageMetadata.InspectExifProfile(profile.Take(profile.Length - 1).ToArray());
        Assert.Contains("exif:invalid-profile", truncated.Issues);
        OfficeImageMetadataInspection limited = OfficeImageMetadata.InspectExifProfile(profile, new OfficeImageMetadataInspectionOptions { MaximumRetainedExifBytes = 7 });
        Assert.Equal(2, limited.ExifFields.Count);
        Assert.Contains(limited.ExifFields, field => field.Omission == "exif:aggregate-value-byte-limit");
    }

    private sealed class ExifField {
        internal OfficeExifDirectory Directory;
        internal ushort Id, Type;
        internal byte[] Bytes = Array.Empty<byte>();
    }
    private static ExifField Field(OfficeExifDirectory directory, ushort id, ushort type, byte[] bytes) => new() { Directory = directory, Id = id, Type = type, Bytes = bytes };
    private static byte[] Ascii(string value) => Encoding.ASCII.GetBytes(value + "\0");
    private static byte[] Rationals(bool little, params (uint Numerator, uint Denominator)[] values) {
        var bytes = new byte[values.Length * 8];
        for (int i = 0; i < values.Length; i++) { Put(bytes, i * 8, values[i].Numerator, 4, little); Put(bytes, i * 8 + 4, values[i].Denominator, 4, little); }
        return bytes;
    }
    // Construct the on-disk contract directly; no OfficeIMO writer participates in these fixtures.
    private static byte[] Exif(bool little, params ExifField[] fields) {
        var groups = fields.GroupBy(field => field.Directory).ToDictionary(group => group.Key, group => group.ToList());
        if (!groups.ContainsKey(OfficeExifDirectory.Image)) groups.Add(OfficeExifDirectory.Image, new List<ExifField>());
        foreach (OfficeExifDirectory directory in new[] { OfficeExifDirectory.Exif, OfficeExifDirectory.Gps }) if (groups.ContainsKey(directory)) groups[OfficeExifDirectory.Image].Add(Field(OfficeExifDirectory.Image, directory == OfficeExifDirectory.Exif ? (ushort)34665 : (ushort)34853, 4, new byte[4]));
        var offsets = new Dictionary<OfficeExifDirectory, int>(); int size = 8;
        foreach (var group in groups.OrderBy(group => group.Key)) { offsets.Add(group.Key, size); size += 6 + group.Value.Count * 12; }
        int payload = size; size += fields.Sum(field => field.Bytes.Length > 4 ? field.Bytes.Length : 0);
        var bytes = new byte[size]; bytes[0] = bytes[1] = little ? (byte)73 : (byte)77; Put(bytes, 2, 42, 2, little); Put(bytes, 4, (uint)offsets[OfficeExifDirectory.Image], 4, little);
        foreach (var group in groups) {
            int at = offsets[group.Key]; Put(bytes, at, (uint)group.Value.Count, 2, little);
            for (int i = 0; i < group.Value.Count; i++) {
                ExifField field = group.Value[i]; int entry = at + 2 + i * 12;
                Put(bytes, entry, field.Id, 2, little); Put(bytes, entry + 2, field.Type, 2, little);
                int itemSize = field.Type == 5 ? 8 : field.Type == 4 ? 4 : 1;
                Put(bytes, entry + 4, (uint)(field.Bytes.Length / itemSize), 4, little);
                if (field.Id == 34665 || field.Id == 34853) Put(bytes, entry + 8, (uint)offsets[field.Id == 34665 ? OfficeExifDirectory.Exif : OfficeExifDirectory.Gps], 4, little);
                else if (field.Bytes.Length <= 4) Buffer.BlockCopy(field.Bytes, 0, bytes, entry + 8, field.Bytes.Length);
                else { Put(bytes, entry + 8, (uint)payload, 4, little); Buffer.BlockCopy(field.Bytes, 0, bytes, payload, field.Bytes.Length); payload += field.Bytes.Length; }
            }
        }
        return bytes;
    }
    private static void Put(byte[] bytes, int offset, uint value, int count, bool little = false) {
        for (int i = 0; i < count; i++) bytes[offset + i] = (byte)(value >> (8 * (little ? i : count - 1 - i)));
    }
}
