using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(0, 10, 20, 5, 0, 25)]
    [InlineData(1, 3, 5, 2, 0, 8)]
    [InlineData(2, 5, 2, 0, 0, 3)]
    [InlineData(0x2003, 0x0147, 0, 0, 0, 25)]
    [InlineData(4, 10, 20, 0, 0, 10)]
    [InlineData(5, 10, 20, 0, 0, 20)]
    [InlineData(6, 0, 10, 20, 0, 20)]
    [InlineData(6, 1, 10, 20, 0, 10)]
    [InlineData(7, 3, 4, 12, 0, 13)]
    [InlineData(8, 0, 1, 0, 0, 5898240)]
    [InlineData(0x4009, 100, 0x0148, 0, 5898240, 100)]
    [InlineData(0x400A, 100, 0x0148, 0, 5898240, 0)]
    [InlineData(11, 100, 3, 4, 0, 60)]
    [InlineData(12, 100, 3, 4, 0, 80)]
    [InlineData(13, 10, 0, 0, 0, 3)]
    [InlineData(14, 0, 1, 2, 0, 196608)]
    [InlineData(15, 3, 5, 10, 0, 8)]
    [InlineData(0x4010, 100, 0x0148, 0, 2949120, 100)]
    public void OfficeArtGeometryGuides_StoredFormulasHaveIntegerAndFixedDegreeSemantics(
        ushort operation, ushort first, ushort second, ushort third, int angle, int expected) {
        var properties = Guides((operation, first, second, third));
        properties.Add(new OfficeArtProperty(1, 0x0147, unchecked((uint)-25)));
        properties.Add(new OfficeArtProperty(2, 0x0148, unchecked((uint)angle)));
        Assert.True(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, _ => { }, default,
            out int[] result, out var failure));
        Assert.Equal(OfficeArtCustomPathFailure.None, failure);
        Assert.Equal(expected, Assert.Single(result));
    }

    [Theory]
    [InlineData(1, -3, 1, 2, -1)]
    [InlineData(1, -3, 1, -2, 2)]
    [InlineData(2, -3, 0, 0, -1)]
    [InlineData(1, 2147483647, 2147483646, 2147483647, 2147483646)]
    public void OfficeArtGeometryGuides_ProductRoundingAndMidpointsKeepSignedPrecision(
        ushort operation, int first, int second, int third, int expected) {
        var properties = Guides(((ushort)(operation | 0xE000), (ushort)0x0147, (ushort)0x0148, (ushort)0x0149));
        properties.Add(new OfficeArtProperty(1, 0x0147, unchecked((uint)first)));
        properties.Add(new OfficeArtProperty(2, 0x0148, unchecked((uint)second)));
        properties.Add(new OfficeArtProperty(3, 0x0149, unchecked((uint)third)));
        Assert.True(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, _ => { }, default, out var result, out _));
        Assert.Equal(expected, Assert.Single(result));
    }

    [Theory]
    [InlineData(0x0140, 60)]
    [InlineData(0x0141, 30)]
    [InlineData(0x0142, 100)]
    [InlineData(0x0143, 100)]
    [InlineData(0x0147, -7)]
    [InlineData(0x0148, -8)]
    [InlineData(0x0149, -9)]
    [InlineData(0x014A, -10)]
    [InlineData(0x014B, -11)]
    [InlineData(0x014C, -12)]
    [InlineData(0x014D, -13)]
    [InlineData(0x014E, -14)]
    [InlineData(0x0153, int.MinValue)]
    [InlineData(0x0154, int.MinValue)]
    [InlineData(0x01FC, 1)]
    [InlineData(0x04FC, 914400)]
    [InlineData(0x04FD, 1828800)]
    [InlineData(0x04FE, 457200)]
    [InlineData(0x04FF, 914400)]
    public void OfficeArtGeometryGuides_OperandsUseGeometrySpaceAndPhysicalFrameUnits(ushort parameter, int expected) {
        var properties = Guides(((ushort)0x2000, parameter, (ushort)0, (ushort)0));
        properties.AddRange(new[] { new OfficeArtProperty(1, 0x0140, 10),
            new OfficeArtProperty(2, 0x0141, unchecked((uint)-20)),
            new OfficeArtProperty(3, 0x0142, 110), new OfficeArtProperty(4, 0x0143, 80) });
        for (ushort id = 0x0147; id <= 0x014E; id++)
            properties.Add(new OfficeArtProperty(properties.Count, id, unchecked((uint)-(id - 0x0140))));
        Assert.True(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, _ => { }, default, out var result, out _));
        Assert.Equal(expected, Assert.Single(result));
    }

    [Theory]
    [InlineData(0U, 1)]
    [InlineData(0x00000008U, 1)]
    [InlineData(0x00080000U, 0)]
    [InlineData(0x00080008U, 1)]
    public void OfficeArtGeometryGuides_StrokeOperandUsesTheNativeUseBitAndDefault(uint flags, int expected) {
        var properties = Guides(((ushort)0x2000, (ushort)0x01FC, (ushort)0, (ushort)0));
        properties.Add(new OfficeArtProperty(1, 0x01FF, flags));
        Assert.True(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, _ => { }, default, out var result, out _));
        Assert.Equal(expected, Assert.Single(result));
    }

    [Fact]
    public void OfficeArtGeometryGuides_PreviousResultsResolveInBothPathAxesAndChargeWorkOnce() {
        var properties = CustomPath(new[] { (unchecked((int)0x80000000), unchecked((int)0x80000001)), (21600, 21600) });
        properties.AddRange(Guides(((ushort)0, (ushort)100, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x0400, (ushort)50, (ushort)0)));
        int items = 0;
        Assert.True(OfficeArtCustomPathProjector.TryProject(properties, 21600, 21600, count => items += count,
            default, out var result, out _));
        Assert.Equal(new OfficePoint(100, 150), result!.Shape.PathCommands[0].Point);
        Assert.True(result.UsesGuides); Assert.Equal(7, items); // Two points, two guides, three commands.
    }

    [Theory]
    [InlineData(0x0400)]
    [InlineData(0x0401)]
    [InlineData(0x047F)]
    public void OfficeArtGeometryGuides_SelfAndForwardReferencesRejectTheWholePath(ushort parameter) {
        var properties = Guides(((ushort)0x2000, parameter, (ushort)0, (ushort)0));
        Assert.False(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, _ => { }, default, out var values, out var failure));
        Assert.Empty(values); Assert.Equal(OfficeArtCustomPathFailure.InvalidGuide, failure);
    }

    [Theory]
    [InlineData(0x04F7)]
    [InlineData(0x04F8)]
    [InlineData(0x04F9)]
    [InlineData(0x0144)]
    public void OfficeArtGeometryGuides_DeviceAndUnknownOperandsRemainUnsupported(ushort parameter) {
        var properties = Guides(((ushort)0x2000, parameter, (ushort)0, (ushort)0));
        Assert.False(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, _ => { }, default, out _, out var failure));
        Assert.Equal(OfficeArtCustomPathFailure.GuideParameter, failure);
    }

    [Theory]
    [InlineData(1, 10, 20, 0)]
    [InlineData(8, 0, 0, 0)]
    [InlineData(11, 1, 0, 0)]
    [InlineData(15, 5, 3, 10)]
    [InlineData(15, 1, 0, 10)]
    public void OfficeArtGeometryGuides_UndefinedArithmeticDoesNotYieldAnExportableCoordinate(
        ushort operation, ushort first, ushort second, ushort third) {
        Assert.False(OfficeArtGeometryGuides.TryEvaluate(Guides((operation, first, second, third)),
            72, 144, _ => { }, default, out var values, out var failure));
        Assert.Empty(values); Assert.Equal(OfficeArtCustomPathFailure.InvalidGuide, failure);
    }

    [Fact]
    public void OfficeArtGeometryGuides_OverflowAndTangentPolesDoNotWrapOrEscapeTheNumericDomain() {
        foreach (ushort operation in new ushort[] { 0x2003, 0x200D }) {
            var properties = Guides((operation, (ushort)0x0147, (ushort)0x0147, (ushort)0));
            properties.Add(new OfficeArtProperty(1, 0x0147, unchecked((uint)int.MinValue)));
            Assert.False(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, _ => { }, default, out _, out _));
        }
        var overflow = Guides(((ushort)0x6001, (ushort)0x0147, (ushort)0x0147, (ushort)1));
        overflow.Add(new OfficeArtProperty(1, 0x0147, int.MaxValue));
        Assert.False(OfficeArtGeometryGuides.TryEvaluate(overflow, 72, 144, _ => { }, default, out _, out _));
        var pole = Guides(((ushort)0x4010, (ushort)100, (ushort)0x0147, (ushort)0));
        pole.Add(new OfficeArtProperty(1, 0x0147, 90 * 65536));
        Assert.False(OfficeArtGeometryGuides.TryEvaluate(pole, 72, 144, _ => { }, default, out _, out _));
    }

    [Fact]
    public void OfficeArtGeometryGuides_RecordLimitAndMalformedArraysRejectBeforeChargingOrAllocatingResults() {
        foreach (byte[] data in new[] { GuideArray(new (ushort, ushort, ushort, ushort)[129]),
            new byte[] { 1, 0, 1, 0, 8, 0 }, new byte[] { 1, 0, 0, 0, 8, 0 }, new byte[] { 0, 0, 0, 0, 8, 0 } }) {
            int items = 0;
            var properties = new[] { new OfficeArtProperty(0, 0x8156, (uint)data.Length, data.Length, complexData: data) };
            Assert.False(OfficeArtGeometryGuides.TryEvaluate(properties, 72, 144, count => items += count,
                default, out var values, out var failure));
            Assert.Empty(values); Assert.Equal(0, items); Assert.Equal(OfficeArtCustomPathFailure.InvalidGuide, failure);
        }
    }

    [Fact]
    public void OfficeArtGeometryGuides_LastGuideAndCumulativeBudgetRespectThe128RecordBoundary() {
        var records = Enumerable.Range(0, 128).Select(index => ((ushort)0, (ushort)index, (ushort)0, (ushort)0)).ToArray();
        var properties = CustomPath(new[] { (unchecked((int)0x8000007F), 0), (21600, 21600) });
        properties.AddRange(Guides(records));
        int charged = 0;
        Assert.True(OfficeArtCustomPathProjector.TryProject(properties, 21600, 21600, count => charged += count,
            default, out var result, out _));
        Assert.Equal(127, result!.Shape.PathCommands[0].Point.X); Assert.Equal(133, charged);
        using var source = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => OfficeArtCustomPathProjector.TryProject(properties,
            21600, 21600, count => { if (count == 128) source.Cancel(); }, source.Token, out _, out _));
        Assert.Throws<InvalidDataException>(() => OfficeArtGeometryGuides.TryEvaluate(properties,
            72, 144, _ => throw new InvalidDataException("work"), default, out _, out _));
    }

    private static List<OfficeArtProperty> Guides(params (ushort Operation, ushort First, ushort Second, ushort Third)[] records) {
        byte[] data = GuideArray(records);
        return new() { new OfficeArtProperty(0, 0x8156, (uint)data.Length, data.Length, complexData: data) };
    }
    private static byte[] GuideArray((ushort Operation, ushort First, ushort Second, ushort Third)[] records) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(checked((ushort)records.Length)); writer.Write(checked((ushort)records.Length)); writer.Write((ushort)8);
        foreach (var record in records) {
            writer.Write(record.Operation); writer.Write(record.First); writer.Write(record.Second); writer.Write(record.Third);
        }
        return stream.ToArray();
    }
}
