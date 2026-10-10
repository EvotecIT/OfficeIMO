using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(4)]
    [InlineData(0xFFF0)]
    public void OfficeArtCustomPath_CompactCoordinatesPreserveTheirUnsignedRange(int elementSize) {
        var properties = CustomPath(new[] { (32767, 32768), (40000, 65535) }, elementSize: elementSize);
        properties.Add(new OfficeArtProperty(1, 0x0142, 65535));
        properties.Add(new OfficeArtProperty(2, 0x0143, 65535));
        Assert.True(OfficeArtCustomPathProjector.TryProject(properties, 65535, 65535, _ => { },
            default, out var result, out _));
        Assert.Equal(new OfficePoint(32767, 32768), result!.Shape.PathCommands[0].Point);
        Assert.Equal(new OfficePoint(40000, 65535), result.Shape.PathCommands[1].Point);
    }

    [Fact]
    public void OfficeArtCustomPath_FullCoordinatesRetainNegativeSignedValues() {
        var properties = CustomPath(new[] { (-40000, -32768), (40000, 65535) });
        Assert.True(OfficeArtCustomPathProjector.TryProject(properties, 21600, 21600, _ => { },
            default, out var result, out _));
        Assert.Equal(new OfficePoint(-40000, -32768), result!.Shape.PathCommands[0].Point);
    }

    [Fact]
    public void OfficeArtCustomPath_CommandRunsConsumeAllVerticesAndPreserveSubpaths() {
        var properties = CustomPath(new[] { (0, 0), (21600, 0), (21600, 21600), (0, 21600),
            (5400, 5400), (10800, 5400), (10800, 10800), (5400, 10800) },
            new ushort[] { 0x4000, 0xAC00, 0x0003, 0x6001, 0x4000, 0xAE00, 0x0003, 0x6001, 0x8000 });
        int items = 0;
        Assert.True(OfficeArtCustomPathProjector.TryProject(properties, 80, 40, count => items += count,
            default, out var result, out _));
        Assert.Equal(27, items); // Eight vertices, nine native commands and ten projected commands.
        Assert.Equal(80, result!.Shape.Width); Assert.Equal(40, result.Shape.Height);
        Assert.Equal(OfficeFillRule.NonZero, result.Shape.FillRule);
        Assert.Equal(2, result.Shape.PathCommands.Count(command => command.Kind == OfficePathCommandKind.MoveTo));
        Assert.Equal(new OfficePoint(20, 10), result.Shape.PathCommands[5].Point);
    }

    [Theory]
    [InlineData(0x02000000U, true, false)]
    [InlineData(0x02000200U, false, false)]
    [InlineData(0x00000200U, false, false)]
    [InlineData(0x00400000U, false, true)]
    [InlineData(0x00400040U, false, false)]
    public void OfficeArtCustomPath_GeometryPaintControlsRequireTheirUseBits(uint flags, bool noFill, bool noLine) {
        var properties = CustomPath(new[] { (0, 0), (21600, 21600) });
        properties.Add(new OfficeArtProperty(properties.Count, 0x017F, flags));
        Assert.True(OfficeArtCustomPathProjector.TryProject(properties, 80, 40, _ => { }, default, out var result, out _));
        Assert.Equal(noFill, result!.NoFill); Assert.Equal(noLine, result.NoLine);
    }

    [Fact]
    public void OfficeArtCustomPath_CubicRunRetainsBothControlsAndEndWithoutClosing() {
        var properties = CustomPath(new[] { (0, 0), (7200, 0), (14400, 21600), (21600, 21600) },
            new ushort[] { 0x4000, 0xAD00, 0x2001, 0xAB00, 0x8000 });
        Assert.True(OfficeArtCustomPathProjector.TryProject(properties, 90, 60, _ => { }, default, out var result, out _));
        Assert.Equal(2, result!.Shape.PathCommands.Count);
        OfficePathCommand curve = result.Shape.PathCommands[1];
        Assert.Equal(new OfficePoint(30, 0), curve.ControlPoint1);
        Assert.Equal(new OfficePoint(60, 60), curve.ControlPoint2);
        Assert.Equal(new OfficePoint(90, 60), curve.Point);
        Assert.True(result.NoLine);
    }

    [Fact]
    public void OfficeArtCustomPath_GuideSentinelsAreNeverInterpretedAsLiteralCoordinates() {
        var properties = CustomPath(new[] { (unchecked((int)0x8000007F), 0), (21600, 21600) });
        Assert.False(OfficeArtCustomPathProjector.TryProject(properties, 80, 40, _ => { }, default, out var result, out var failure));
        Assert.Null(result); Assert.Equal(OfficeArtCustomPathFailure.GuideReference, failure);
    }

    [Theory]
    [InlineData(0x6000)]
    [InlineData(0x4001)]
    [InlineData(0x0003)]
    [InlineData(0x2001)]
    public void OfficeArtCustomPath_InvalidPointConsumptionDoesNotReturnPartialArtwork(ushort command) {
        var properties = CustomPath(new[] { (0, 0), (21600, 21600) }, new ushort[] { 0x4000, command, 0x8000 });
        Assert.False(OfficeArtCustomPathProjector.TryProject(properties, 80, 40, _ => { }, default, out var result, out _));
        Assert.Null(result);
    }

    [Fact]
    public void OfficeArtCustomPath_DeclaredAndExpandedWorkUsesTheImporterBudgetAndCancellation() {
        var properties = CustomPath(new[] { (0, 0), (21600, 21600) });
        int items = 0;
        Assert.Throws<InvalidDataException>(() => OfficeArtCustomPathProjector.TryProject(properties, 80, 40,
            count => { items += count; if (items > 3) throw new InvalidDataException("work"); },
            default, out _, out _));
        using var source = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => OfficeArtCustomPathProjector.TryProject(properties, 80, 40,
            _ => source.Cancel(), source.Token, out _, out _));
    }

    private static List<OfficeArtProperty> CustomPath((int X, int Y)[] points, ushort[]? segments = null, int elementSize = 8) {
        byte[] ArrayData(int count, ushort stride, Action<BinaryWriter> write) {
            using var data = new MemoryStream(); using var writer = new BinaryWriter(data);
            writer.Write(checked((ushort)count)); writer.Write(checked((ushort)count)); writer.Write(stride);
            write(writer); return data.ToArray();
        }
        var properties = new List<OfficeArtProperty>();
        byte[] vertices = ArrayData(points.Length, checked((ushort)elementSize), writer => {
            foreach (var point in points) {
                if (elementSize == 8) { writer.Write(point.X); writer.Write(point.Y); }
                else { writer.Write(checked((ushort)point.X)); writer.Write(checked((ushort)point.Y)); }
            }
        });
        properties.Add(new OfficeArtProperty(0, 0x8145, (uint)vertices.Length, vertices.Length, complexData: vertices));
        if (segments != null) {
            byte[] words = ArrayData(segments.Length, 2, writer => { foreach (ushort word in segments) writer.Write(word); });
            properties.Add(new OfficeArtProperty(1, 0x8146, (uint)words.Length, words.Length, complexData: words));
        }
        return properties;
    }
}
