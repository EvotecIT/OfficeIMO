using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfToUnicodeRetentionTests {
    [Fact]
    public void AcceptedLongDestinationEstimateCoversItsRetainedUtf16Payload() {
        string destination = new string('A', 4096);
        ToUnicodeCMap cmap = ParseDestinations(destination);

        Assert.Equal(destination, cmap.MapBytes(new byte[] { 1 }, destination.Length));
        Assert.True(cmap.EstimatedRetainedBytes >= Encoding.Unicode.GetByteCount(destination));
    }

    [Fact]
    public void ReplacedForwardMappingStillChargesItsUsableReverseHistory() {
        string[] destinations = Enumerable.Range(0x100, 64).Select(value => new string((char)value, 64)).ToArray();
        ToUnicodeCMap cmap = ParseDestinations(destinations);

        Assert.Equal(destinations.Last(), cmap.MapBytes(new byte[] { 1 }));
        Assert.True(cmap.TryEncodeText(destinations.First(), out string firstCode));
        Assert.Equal("01", firstCode);
        Assert.True(cmap.TryEncodeText(destinations.Last(), out string lastCode));
        Assert.Equal("01", lastCode);
        Assert.True(cmap.EstimatedRetainedBytes > destinations.Sum(destination => Encoding.Unicode.GetByteCount(destination)));
        Assert.True(cmap.EstimatedRetainedBytes > ParseDestinations(destinations.Last()).EstimatedRetainedBytes);
    }

    [Fact]
    public void SupersededNonReversibleDestinationStopsConsumingRetainedEstimate() {
        string original = new string('A', 4096);
        string replacement = new string('B', 65);
        ToUnicodeCMap overwritten = ParseDestinations(original, replacement);
        ToUnicodeCMap finalOnly = ParseDestinations(replacement);

        Assert.Equal(replacement, overwritten.MapBytes(new byte[] { 1 }));
        Assert.Equal(replacement, finalOnly.MapBytes(new byte[] { 1 }));
        Assert.False(overwritten.TryEncodeText(original, out _));
        Assert.False(overwritten.TryEncodeText(replacement, out _));
        Assert.Equal(finalOnly.EstimatedRetainedBytes, overwritten.EstimatedRetainedBytes);
    }

    private static ToUnicodeCMap ParseDestinations(params string[] destinations) {
        var source = new StringBuilder("beginbfchar\n");
        foreach (string destination in destinations) {
            source.Append("<01> <");
            foreach (char character in destination) source.Append(((int)character).ToString("X4", System.Globalization.CultureInfo.InvariantCulture));
            source.Append(">\n");
        }
        source.Append("endbfchar");
        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source.ToString()), out ToUnicodeCMap? cmap));
        return Assert.IsType<ToUnicodeCMap>(cmap);
    }
}
