using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffExternalWritableRangesTests {
    [Fact]
    public void BatchValidationRejectsAliasedValuesAndDirectoryStorage() {
        (byte[] tiff, int[] ranges) = CreateTiff(3);
        Assert.True(OfficeTiffStructureValidator.TryValidateExclusiveExternalWritableRanges(
            tiff, 0, tiff.Length, CancellationToken.None, ranges));

        byte[] aliased = (byte[])tiff.Clone();
        WriteUInt32(aliased, Entry(1) + 8, (uint)ranges[0]);
        int[] aliasedRanges = (int[])ranges.Clone();
        aliasedRanges[3] = ranges[0];
        Assert.True(OfficeTiffStructureValidator.TryValidate(aliased, 0, aliased.Length));
        Assert.False(OfficeTiffStructureValidator.TryValidateExclusiveExternalWritableRanges(
            aliased, 0, aliased.Length, CancellationToken.None, aliasedRanges));

        byte[] tableOwned = (byte[])tiff.Clone();
        WriteUInt32(tableOwned, Entry(0) + 8, 8);
        int[] tableRanges = (int[])ranges.Clone();
        tableRanges[0] = 8;
        Assert.True(OfficeTiffStructureValidator.TryValidate(tableOwned, 0, tableOwned.Length));
        Assert.False(OfficeTiffStructureValidator.TryValidateExclusiveExternalWritableRanges(
            tableOwned, 0, tableOwned.Length, CancellationToken.None, tableRanges));
    }

    [Fact]
    [Trait("Category", "Performance")]
    public void BatchValidationMeasuresManyIndependentExifValues() {
        (byte[] tiff, int[] ranges) = CreateTiff(2000);
        var repeated = Stopwatch.StartNew();
        for (int index = 0; index < ranges.Length; index += 3) {
            Assert.True(OfficeTiffStructureValidator.TryValidateExclusiveWritableRanges(
                tiff, 0, tiff.Length, CancellationToken.None,
                new[] { ranges[index], ranges[index + 1], ranges[index + 2] }));
        }
        repeated.Stop();
        var batch = Stopwatch.StartNew();
        Assert.True(OfficeTiffStructureValidator.TryValidateExclusiveExternalWritableRanges(
            tiff, 0, tiff.Length, CancellationToken.None, ranges));
        batch.Stop();
        Console.WriteLine($"TIFF EXIF 2000 fields: repeated {repeated.Elapsed.TotalMilliseconds:F2} ms; batch {batch.Elapsed.TotalMilliseconds:F2} ms.");
    }

    private static (byte[] Tiff, int[] Ranges) CreateTiff(int count) {
        int values = checked(8 + 2 + count * 12 + 4);
        var bytes = new byte[checked(values + count * 8)];
        bytes[0] = (byte)'I'; bytes[1] = (byte)'I';
        WriteUInt16(bytes, 2, 42);
        WriteUInt32(bytes, 4, 8);
        WriteUInt16(bytes, 8, checked((ushort)count));
        var ranges = new List<int>(count * 3);
        for (int index = 0; index < count; index++) {
            int entry = Entry(index), value = values + index * 8;
            WriteUInt16(bytes, entry, checked((ushort)(1000 + index)));
            WriteUInt16(bytes, entry + 2, 7); // UNDEFINED, eight external bytes.
            WriteUInt32(bytes, entry + 4, 8);
            WriteUInt32(bytes, entry + 8, (uint)value);
            ranges.Add(value); ranges.Add(8); ranges.Add(entry);
        }
        return (bytes, ranges.ToArray());
    }

    private static int Entry(int index) => 10 + index * 12;
    private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)value; bytes[offset + 1] = (byte)(value >> 8);
    }
    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value; bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16); bytes[offset + 3] = (byte)(value >> 24);
    }
}
