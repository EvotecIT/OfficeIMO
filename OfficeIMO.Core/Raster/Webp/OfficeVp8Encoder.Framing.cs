// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private byte[] BuildKeyframeHeader(int width, int height) {
        var header = new byte[7];
        header[0] = 0x9D;
        header[1] = 0x01;
        header[2] = 0x2A;
        WriteU16LE(header, 3, width & 0x3FFF);
        WriteU16LE(header, 5, height & 0x3FFF);
        return header;
    }

    private byte[] BuildFrameTag(int partitionSize, int version, bool showFrame, bool keyframe) {
        var tag = (partitionSize << 5) | ((showFrame ? 1 : 0) << 4) | (version << 1) | (keyframe ? 0 : 1);
        var bytes = new byte[3];
        bytes[0] = (byte)(tag & 0xFF);
        bytes[1] = (byte)((tag >> 8) & 0xFF);
        bytes[2] = (byte)((tag >> 16) & 0xFF);
        return bytes;
    }

    private void WriteU16LE(byte[] buffer, int offset, int value) {
        buffer[offset] = (byte)(value & 0xFF);
        buffer[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private void SwapRowContexts(ref byte[] above, ref byte[] current) {
        var temp = above;
        above = current;
        current = temp;
        Array.Clear(current, 0, current.Length);
    }

    private void SwapRowModes(ref int[] above, ref int[] current) {
        var temp = above;
        above = current;
        current = temp;
        Array.Clear(current, 0, current.Length);
    }

    private byte[] PadToLength(byte[] data, int length) {
        if (data.Length >= length) return data;
        var padded = new byte[length];
        Buffer.BlockCopy(data, 0, padded, 0, data.Length);
        return padded;
    }

    private readonly struct DequantFactors {
        public DequantFactors(int y1Dc, int y1Ac, int y2Dc, int y2Ac, int uvDc, int uvAc) {
            Y1Dc = y1Dc;
            Y1Ac = y1Ac;
            Y2Dc = y2Dc;
            Y2Ac = y2Ac;
            UvDc = uvDc;
            UvAc = uvAc;
        }

        public int Y1Dc { get; }
        public int Y1Ac { get; }
        public int Y2Dc { get; }
        public int Y2Ac { get; }
        public int UvDc { get; }
        public int UvAc { get; }
    }
}
