using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingHuffmanTests {
    [Theory]
    [InlineData(1, 1, false)]
    [InlineData(19, 13, false)]
    [InlineData(1, 1, true)]
    [InlineData(19, 13, true)]
    public void OptimizedJpegTablesReserveTheAllOnePaddingCode(int width, int height, bool progressive) {
        var random = new Random(2720);
        var pixels = new byte[width * height * 4];
        for (int index = 0; index < pixels.Length; index += 4) {
            pixels[index] = (byte)random.Next(256);
            pixels[index + 1] = (byte)random.Next(256);
            pixels[index + 2] = (byte)random.Next(256);
            pixels[index + 3] = 255;
        }
        byte[] jpeg = OfficeJpegWriter.WriteRgba(width, height, pixels, width * 4,
            new OfficeJpegEncodeOptions { Quality = 85, Subsampling = OfficeJpegSubsampling.Y444, OptimizeHuffman = true, Progressive = progressive });
        int tables = 0;
        for (int offset = 2; offset + 4 <= jpeg.Length;) {
            Assert.Equal(255, jpeg[offset]);
            if (jpeg[offset + 1] == 0xDA) break;
            int end = offset + 2 + (jpeg[offset + 2] << 8 | jpeg[offset + 3]);
            if (jpeg[offset + 1] == 0xC4) {
                for (int position = offset + 4; position < end;) {
                    position++; // table identity
                    int codeSpace = 0, symbols = 0;
                    for (int depth = 1; depth <= 16; depth++) {
                        int count = jpeg[position++];
                        symbols += count;
                        codeSpace += count * (1 << (16 - depth));
                    }
                    Assert.InRange(codeSpace, 1, (1 << 16) - 1);
                    position += symbols;
                    tables++;
                }
            }
            offset = end;
        }
        Assert.True(tables > 0);
    }

    [Theory]
    [InlineData(15)]
    [InlineData(16)]
    public void SkewedHuffmanHistogramTerminatesWithACompleteDepthLimitedTree(int maximumDepth) {
        var frequencies = new int[20];
        frequencies[0] = frequencies[1] = 1;
        for (int index = 2; index < frequencies.Length; index++) frequencies[index] = frequencies[index - 1] + frequencies[index - 2];
        byte[] lengths = OfficeHuffmanCodeLengths.Create(frequencies, maximumDepth);
        Assert.All(lengths, length => Assert.InRange(length, (byte)1, (byte)maximumDepth));
        Assert.Equal(1L << maximumDepth, lengths.Sum(length => 1L << (maximumDepth - length)));
        Assert.True(lengths[0] > lengths[lengths.Length - 1]);
        Assert.Equal(lengths, OfficeHuffmanCodeLengths.Create(frequencies, maximumDepth));
    }

    [Theory]
    [InlineData(280)]
    [InlineData(256)]
    [InlineData(40)]
    public void HuffmanHistogramPreservesSparseAlphabetAndFavorsFrequentSymbols(int alphabet) {
        var frequencies = new int[alphabet];
        for (int symbol = 0; symbol < alphabet; symbol += 3) frequencies[symbol] = symbol == 0 ? 10000 : 1 + symbol % 7;
        byte[] lengths = OfficeHuffmanCodeLengths.Create(frequencies, 15);
        for (int symbol = 0; symbol < alphabet; symbol++) {
            if (frequencies[symbol] == 0) Assert.Equal(0, lengths[symbol]);
            else Assert.InRange(lengths[symbol], (byte)1, (byte)15);
        }
        Assert.Equal(1L << 15, lengths.Where(length => length > 0).Sum(length => 1L << (15 - length)));
        Assert.Equal(1, lengths[0]);
    }
}
