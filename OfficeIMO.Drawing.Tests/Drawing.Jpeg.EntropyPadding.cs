using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegEntropyPaddingTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void ZeroCoefficientScansUseOneFillBits(bool progressive, bool optimized, bool color) {
        byte[] jpeg = OfficeJpegWriter.WriteRgba(1, 1, new byte[] { 128, 128, color ? (byte)129 : (byte)128, 255 }, 4,
            new OfficeJpegEncodeOptions { Quality = color ? 1 : 85, Subsampling = OfficeJpegSubsampling.Y444, Progressive = progressive, OptimizeHuffman = optimized });
        var scans = new List<byte[]>();
        for (int offset = 2; offset + 4 <= jpeg.Length;) {
            Assert.Equal(255, jpeg[offset]);
            int marker = jpeg[offset + 1];
            if (marker == 0xD9) break;
            int end = offset + 2 + (jpeg[offset + 2] << 8 | jpeg[offset + 3]);
            if (marker == 0xDA) {
                var bytes = new List<byte>();
                while (end < jpeg.Length) {
                    byte value = jpeg[end++];
                    if (value == 255) {
                        if (jpeg[end] != 0) { end--; break; }
                        end++;
                    }
                    bytes.Add(value);
                }
                scans.Add(bytes.ToArray());
            }
            offset = end;
        }

        // Neutral grey uses one component. The near-grey RGB fixture quantizes
        // to zero at quality 1 while preserving three components. Both have
        // zero DC differences and only EOB AC symbols, so the complete scan
        // bits are known independently of the encoder's bit-writing helper.
        if (!progressive) {
            Assert.Single(scans);
            byte[] expected = color
                ? optimized ? new byte[] { 0x03 } : new byte[] { 0x28, 0x03 }
                : new byte[] { optimized ? (byte)0x3F : (byte)0x2B };
            Assert.Equal(expected, scans[0]);
        } else {
            Assert.Equal(color ? 4 : 2, scans.Count);
            byte dc = color ? optimized ? (byte)0x1F : (byte)0x03 : optimized ? (byte)0x7F : (byte)0x3F;
            Assert.Equal(new byte[] { dc }, scans[0]);
            Assert.Equal(new byte[] { optimized ? (byte)0x7F : (byte)0xAF }, scans[1]);
            if (color) {
                Assert.Equal(new byte[] { optimized ? (byte)0x7F : (byte)0x3F }, scans[2]);
                Assert.Equal(scans[2], scans[3]);
            }
        }
    }
}
