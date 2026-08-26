using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal static partial class OfficeWoff2Decoder {
    private const int TransformedGlyfHeaderLength = 36;
    private const ushort CompositeArgumentsAreWords = 0x0001;
    private const ushort CompositeHasScale = 0x0008;
    private const ushort CompositeMoreComponents = 0x0020;
    private const ushort CompositeHasXYScale = 0x0040;
    private const ushort CompositeHasTwoByTwo = 0x0080;
    private const ushort CompositeHasInstructions = 0x0100;

    private static GlyfResult ReconstructGlyf(
        byte[] data,
        int expectedGlyfLength,
        int expectedLocaLength,
        int maximumDecodedBytes) {
        if (data.Length < TransformedGlyfHeaderLength) throw new InvalidDataException("The transformed WOFF 2 glyf header is truncated.");
        ushort version = ReadUInt16(data, 0);
        ushort optionFlags = ReadUInt16(data, 2);
        int glyphCount = ReadUInt16(data, 4);
        ushort indexFormat = ReadUInt16(data, 6);
        if (version != 0) throw new NotSupportedException("The WOFF 2 glyf transform version is not supported.");
        if ((optionFlags & ~1) != 0) throw new InvalidDataException("The transformed WOFF 2 glyf option flags are invalid.");
        if (glyphCount <= 0) throw new InvalidDataException("The transformed WOFF 2 glyf table has no glyphs.");
        if (indexFormat > 1) throw new InvalidDataException("The transformed WOFF 2 loca index format is invalid.");
        int locaEntrySize = indexFormat == 0 ? 2 : 4;
        int requiredLocaLength = checked((glyphCount + 1) * locaEntrySize);
        if (expectedGlyfLength <= 0 || expectedLocaLength != requiredLocaLength) {
            throw new InvalidDataException("The transformed WOFF 2 glyf/loca output lengths are invalid.");
        }
        // The directory's original glyf length is reference metadata for transformed tables;
        // conforming fonts can reconstruct to a different byte count. Bound the actual builder
        // and all per-glyph builders independently instead of requiring byte-count equality.
        long retainedReconstructionBytes = checked(
            data.LongLength * 2L +
            expectedLocaLength + 24L +
            (glyphCount + 1L) * sizeof(uint) + 24L +
            glyphCount * sizeof(short) + 24L +
            8L * 24L);
        long availableGlyfWorkingBytes = maximumDecodedBytes - retainedReconstructionBytes;
        if (availableGlyfWorkingBytes < 3L) {
            throw new InvalidDataException("The reconstructed WOFF 2 glyf working set exceeds the configured byte limit.");
        }
        int maximumGlyfBytes = checked((int)Math.Min(int.MaxValue, availableGlyfWorkingBytes / 8L));

        var sizes = new int[7];
        int headerCursor = 8;
        long streamsLength = TransformedGlyfHeaderLength;
        for (int index = 0; index < sizes.Length; index++) {
            uint value = ReadUInt32(data, headerCursor);
            headerCursor += 4;
            if (value > int.MaxValue) throw new InvalidDataException("A transformed WOFF 2 glyf stream is too large.");
            sizes[index] = checked((int)value);
            streamsLength = checked(streamsLength + value);
        }
        int overlapBitmapLength = (optionFlags & 1) != 0 ? checked((glyphCount + 7) / 8) : 0;
        if (streamsLength + overlapBitmapLength != data.Length) {
            throw new InvalidDataException("The transformed WOFF 2 glyf stream sizes are inconsistent.");
        }

        int cursor = TransformedGlyfHeaderLength;
        byte[] nContourStream = Slice(data, ref cursor, sizes[0]);
        byte[] nPointsStream = Slice(data, ref cursor, sizes[1]);
        byte[] flagStream = Slice(data, ref cursor, sizes[2]);
        byte[] glyphStream = Slice(data, ref cursor, sizes[3]);
        byte[] compositeStream = Slice(data, ref cursor, sizes[4]);
        byte[] bboxStream = Slice(data, ref cursor, sizes[5]);
        byte[] instructionStream = Slice(data, ref cursor, sizes[6]);
        byte[] overlapBitmap = overlapBitmapLength == 0 ? Array.Empty<byte>() : Slice(data, ref cursor, overlapBitmapLength);
        if (nContourStream.Length != glyphCount * 2) {
            throw new InvalidDataException("The transformed WOFF 2 contour stream length is invalid.");
        }
        int bboxBitmapLength = checked(((glyphCount + 31) >> 5) << 2);
        if (bboxStream.Length < bboxBitmapLength) throw new InvalidDataException("The transformed WOFF 2 bounding-box bitmap is truncated.");
        var bboxBitmap = new byte[bboxBitmapLength];
        Buffer.BlockCopy(bboxStream, 0, bboxBitmap, 0, bboxBitmap.Length);
        int bboxCursor = bboxBitmapLength;

        int pointsCursor = 0;
        int flagsCursor = 0;
        int glyphCursor = 0;
        int compositeCursor = 0;
        int instructionCursor = 0;
        var glyf = new ByteBuilder(maximumGlyfBytes);
        var offsets = new uint[glyphCount + 1];
        var xMinimums = new short[glyphCount];
        for (int glyphId = 0; glyphId < glyphCount; glyphId++) {
            offsets[glyphId] = checked((uint)glyf.Count);
            short contourCount = ReadInt16(nContourStream, glyphId * 2);
            byte[] glyph;
            if (contourCount == 0) {
                glyph = Array.Empty<byte>();
            } else {
                int maximumGlyphBytes = glyf.RemainingCapacity;
                bool hasBox = (bboxBitmap[glyphId >> 3] & (0x80 >> (glyphId & 7))) != 0;
                if (contourCount > 0) {
                    glyph = DecodeSimpleGlyph(
                    glyphId,
                    contourCount,
                    hasBox,
                    bboxStream,
                    ref bboxCursor,
                    nPointsStream,
                    ref pointsCursor,
                    flagStream,
                    ref flagsCursor,
                    glyphStream,
                    ref glyphCursor,
                    instructionStream,
                    ref instructionCursor,
                    overlapBitmap,
                    maximumGlyphBytes,
                    out short xMin);
                    xMinimums[glyphId] = xMin;
                } else if (contourCount == -1) {
                    if (!hasBox) throw new InvalidDataException("A transformed WOFF 2 composite glyph is missing its bounding box.");
                    glyph = DecodeCompositeGlyph(
                        bboxStream,
                        ref bboxCursor,
                        glyphStream,
                        ref glyphCursor,
                        compositeStream,
                        ref compositeCursor,
                        instructionStream,
                        ref instructionCursor,
                        maximumGlyphBytes,
                        out short xMin);
                    xMinimums[glyphId] = xMin;
                } else {
                    throw new InvalidDataException("A transformed WOFF 2 glyph has an invalid contour count.");
                }
            }
            glyf.Add(glyph);
            glyf.PadToEven();
        }
        offsets[glyphCount] = checked((uint)glyf.Count);
        if (pointsCursor != nPointsStream.Length
            || flagsCursor != flagStream.Length
            || glyphCursor != glyphStream.Length
            || compositeCursor != compositeStream.Length
            || bboxCursor != bboxStream.Length
            || instructionCursor != instructionStream.Length) {
            throw new InvalidDataException("A transformed WOFF 2 glyf substream contains trailing or missing data.");
        }

        var loca = new byte[expectedLocaLength];
        int locaOffset = 0;
        if (indexFormat == 0) {
            foreach (uint offsetValue in offsets) {
                if ((offsetValue & 1) != 0 || offsetValue / 2 > ushort.MaxValue) {
                    throw new InvalidDataException("The reconstructed WOFF 2 loca table cannot use short offsets.");
                }
                WriteUInt16(loca, locaOffset, checked((ushort)(offsetValue / 2)));
                locaOffset += 2;
            }
        } else {
            foreach (uint offsetValue in offsets) {
                WriteUInt32(loca, locaOffset, offsetValue);
                locaOffset += 4;
            }
        }
        return new GlyfResult(glyf.ToArray(), loca, indexFormat, xMinimums);
    }

    private static byte[] DecodeSimpleGlyph(
        int glyphId,
        int contourCount,
        bool hasBox,
        byte[] bboxStream,
        ref int bboxCursor,
        byte[] pointsStream,
        ref int pointsCursor,
        byte[] flagStream,
        ref int flagCursor,
        byte[] glyphStream,
        ref int glyphCursor,
        byte[] instructionStream,
        ref int instructionCursor,
        byte[] overlapBitmap,
        int maximumGlyphBytes,
        out short xMinimum) {
        var endPoints = new ushort[contourCount];
        int pointCount = 0;
        for (int contour = 0; contour < contourCount; contour++) {
            int contourPoints = Read255UInt16(pointsStream, ref pointsCursor);
            if (contourPoints <= 0 || pointCount > ushort.MaxValue - contourPoints) {
                throw new InvalidDataException("A transformed WOFF 2 simple glyph has an invalid point count.");
            }
            pointCount += contourPoints;
            endPoints[contour] = checked((ushort)(pointCount - 1));
        }
        if (pointCount > maximumGlyphBytes / 2) throw new InvalidDataException("A transformed WOFF 2 glyph exceeds the configured point budget.");
        EnsureAvailable(flagStream, flagCursor, pointCount, "The transformed WOFF 2 point-flag stream is truncated.");
        var points = new GlyphPoint[pointCount];
        int x = 0;
        int y = 0;
        for (int pointIndex = 0; pointIndex < pointCount; pointIndex++) {
            byte rawFlag = flagStream[flagCursor++];
            bool onCurve = (rawFlag & 0x80) == 0;
            int tripletFlag = rawFlag & 0x7F;
            int byteCount = tripletFlag < 84 ? 1 : tripletFlag < 120 ? 2 : tripletFlag < 124 ? 3 : 4;
            EnsureAvailable(glyphStream, glyphCursor, byteCount, "The transformed WOFF 2 coordinate stream is truncated.");
            DecodeTriplet(glyphStream, ref glyphCursor, tripletFlag, out int deltaX, out int deltaY);
            x = checked(x + deltaX);
            y = checked(y + deltaY);
            if (x < short.MinValue || x > short.MaxValue || y < short.MinValue || y > short.MaxValue) {
                throw new InvalidDataException("A transformed WOFF 2 glyph coordinate exceeds the TrueType range.");
            }
            points[pointIndex] = new GlyphPoint((short)x, (short)y, onCurve);
        }

        int instructionLength = Read255UInt16(glyphStream, ref glyphCursor);
        EnsureAvailable(instructionStream, instructionCursor, instructionLength, "The transformed WOFF 2 instruction stream is truncated.");
        var instructions = new byte[instructionLength];
        Buffer.BlockCopy(instructionStream, instructionCursor, instructions, 0, instructions.Length);
        instructionCursor += instructions.Length;

        short xMin;
        short yMin;
        short xMax;
        short yMax;
        if (hasBox) {
            ReadBoundingBox(bboxStream, ref bboxCursor, out xMin, out yMin, out xMax, out yMax);
        } else {
            CalculateBoundingBox(points, out xMin, out yMin, out xMax, out yMax);
        }
        xMinimum = xMin;

        bool overlapSimple = overlapBitmap.Length > 0
            && (overlapBitmap[glyphId >> 3] & (0x80 >> (glyphId & 7))) != 0;
        var result = new ByteBuilder(maximumGlyphBytes);
        result.AddInt16(checked((short)contourCount));
        result.AddInt16(xMin);
        result.AddInt16(yMin);
        result.AddInt16(xMax);
        result.AddInt16(yMax);
        foreach (ushort endPoint in endPoints) result.AddUInt16(endPoint);
        result.AddUInt16(checked((ushort)instructions.Length));
        result.Add(instructions);
        EncodeSimpleCoordinates(points, overlapSimple, result, maximumGlyphBytes);
        return result.ToArray();
    }

    private static byte[] DecodeCompositeGlyph(
        byte[] bboxStream,
        ref int bboxCursor,
        byte[] glyphStream,
        ref int glyphCursor,
        byte[] compositeStream,
        ref int compositeCursor,
        byte[] instructionStream,
        ref int instructionCursor,
        int maximumGlyphBytes,
        out short xMinimum) {
        ReadBoundingBox(bboxStream, ref bboxCursor, out short xMin, out short yMin, out short xMax, out short yMax);
        xMinimum = xMin;
        var result = new ByteBuilder(maximumGlyphBytes);
        result.AddInt16(-1);
        result.AddInt16(xMin);
        result.AddInt16(yMin);
        result.AddInt16(xMax);
        result.AddInt16(yMax);
        bool moreComponents;
        bool hasInstructions = false;
        do {
            EnsureAvailable(compositeStream, compositeCursor, 4, "The transformed WOFF 2 composite stream is truncated.");
            ushort flags = ReadUInt16(compositeStream, compositeCursor);
            int componentLength = 4;
            componentLength += (flags & CompositeArgumentsAreWords) != 0 ? 4 : 2;
            if ((flags & CompositeHasScale) != 0) componentLength += 2;
            else if ((flags & CompositeHasXYScale) != 0) componentLength += 4;
            else if ((flags & CompositeHasTwoByTwo) != 0) componentLength += 8;
            EnsureAvailable(compositeStream, compositeCursor, componentLength, "The transformed WOFF 2 composite component is truncated.");
            result.Add(compositeStream, compositeCursor, componentLength);
            compositeCursor += componentLength;
            moreComponents = (flags & CompositeMoreComponents) != 0;
            hasInstructions |= (flags & CompositeHasInstructions) != 0;
        } while (moreComponents);

        if (hasInstructions) {
            int instructionLength = Read255UInt16(glyphStream, ref glyphCursor);
            EnsureAvailable(instructionStream, instructionCursor, instructionLength, "The transformed WOFF 2 composite instruction stream is truncated.");
            result.AddUInt16(checked((ushort)instructionLength));
            result.Add(instructionStream, instructionCursor, instructionLength);
            instructionCursor += instructionLength;
        }
        return result.ToArray();
    }

    private static void DecodeTriplet(byte[] data, ref int offset, int flag, out int deltaX, out int deltaY) {
        static int Signed(int signFlags, int value) => (signFlags & 1) != 0 ? value : -value;
        if (flag < 10) {
            deltaX = 0;
            deltaY = Signed(flag, ((flag & 14) << 7) + data[offset]);
            offset += 1;
        } else if (flag < 20) {
            deltaX = Signed(flag, (((flag - 10) & 14) << 7) + data[offset]);
            deltaY = 0;
            offset += 1;
        } else if (flag < 84) {
            int first = flag - 20;
            int value = data[offset++];
            deltaX = Signed(flag, 1 + (first & 0x30) + (value >> 4));
            deltaY = Signed(flag >> 1, 1 + ((first & 0x0C) << 2) + (value & 0x0F));
        } else if (flag < 120) {
            int first = flag - 84;
            deltaX = Signed(flag, 1 + ((first / 12) << 8) + data[offset]);
            deltaY = Signed(flag >> 1, 1 + (((first % 12) >> 2) << 8) + data[offset + 1]);
            offset += 2;
        } else if (flag < 124) {
            int middle = data[offset + 1];
            deltaX = Signed(flag, (data[offset] << 4) + (middle >> 4));
            deltaY = Signed(flag >> 1, ((middle & 0x0F) << 8) + data[offset + 2]);
            offset += 3;
        } else {
            deltaX = Signed(flag, (data[offset] << 8) + data[offset + 1]);
            deltaY = Signed(flag >> 1, (data[offset + 2] << 8) + data[offset + 3]);
            offset += 4;
        }
    }

    private static void EncodeSimpleCoordinates(
        GlyphPoint[] points,
        bool overlapSimple,
        ByteBuilder output,
        int maximumGlyphBytes) {
        var flags = new byte[points.Length];
        var xBytes = new ByteBuilder(maximumGlyphBytes);
        var yBytes = new ByteBuilder(maximumGlyphBytes);
        int previousX = 0;
        int previousY = 0;
        for (int index = 0; index < points.Length; index++) {
            GlyphPoint point = points[index];
            byte flag = point.OnCurve ? (byte)0x01 : (byte)0;
            int deltaX = point.X - previousX;
            int deltaY = point.Y - previousY;
            EncodeCoordinate(deltaX, 0x02, 0x10, ref flag, xBytes);
            EncodeCoordinate(deltaY, 0x04, 0x20, ref flag, yBytes);
            if (index == 0 && overlapSimple) flag |= 0x40;
            flags[index] = flag;
            previousX = point.X;
            previousY = point.Y;
        }
        output.Add(flags);
        output.Add(xBytes.ToArray());
        output.Add(yBytes.ToArray());
    }

    private static void EncodeCoordinate(int delta, byte shortFlag, byte sameOrPositiveFlag, ref byte flags, ByteBuilder output) {
        if (delta == 0) {
            flags |= sameOrPositiveFlag;
        } else if (delta >= -255 && delta <= 255) {
            flags |= shortFlag;
            if (delta > 0) flags |= sameOrPositiveFlag;
            output.Add(checked((byte)Math.Abs(delta)));
        } else {
            if (delta < short.MinValue || delta > short.MaxValue) throw new InvalidDataException("A TrueType coordinate delta exceeds Int16.");
            output.AddUInt16(unchecked((ushort)(short)delta));
        }
    }

    private static void ReadBoundingBox(
        byte[] data,
        ref int offset,
        out short xMin,
        out short yMin,
        out short xMax,
        out short yMax) {
        EnsureAvailable(data, offset, 8, "The transformed WOFF 2 bounding-box stream is truncated.");
        xMin = ReadInt16(data, offset);
        yMin = ReadInt16(data, offset + 2);
        xMax = ReadInt16(data, offset + 4);
        yMax = ReadInt16(data, offset + 6);
        offset += 8;
        if (xMin > xMax || yMin > yMax) throw new InvalidDataException("A transformed WOFF 2 glyph bounding box is invalid.");
    }

    private static void CalculateBoundingBox(
        GlyphPoint[] points,
        out short xMin,
        out short yMin,
        out short xMax,
        out short yMax) {
        if (points.Length == 0) throw new InvalidDataException("A non-empty WOFF 2 glyph contains no points.");
        xMin = xMax = points[0].X;
        yMin = yMax = points[0].Y;
        for (int index = 1; index < points.Length; index++) {
            GlyphPoint point = points[index];
            if (point.X < xMin) xMin = point.X;
            if (point.X > xMax) xMax = point.X;
            if (point.Y < yMin) yMin = point.Y;
            if (point.Y > yMax) yMax = point.Y;
        }
    }

    private static byte[] Slice(byte[] data, ref int offset, int length) {
        EnsureAvailable(data, offset, length, "A transformed WOFF 2 glyf substream is truncated.");
        var result = new byte[length];
        Buffer.BlockCopy(data, offset, result, 0, length);
        offset += length;
        return result;
    }

    private readonly struct GlyfResult {
        internal GlyfResult(byte[] glyf, byte[] loca, ushort indexFormat, short[] xMinimums) {
            Glyf = glyf;
            Loca = loca;
            IndexFormat = indexFormat;
            XMinimums = xMinimums;
        }

        internal byte[] Glyf { get; }
        internal byte[] Loca { get; }
        internal ushort IndexFormat { get; }
        internal short[] XMinimums { get; }
    }

    private readonly struct GlyphPoint {
        internal GlyphPoint(short x, short y, bool onCurve) {
            X = x;
            Y = y;
            OnCurve = onCurve;
        }

        internal short X { get; }
        internal short Y { get; }
        internal bool OnCurve { get; }
    }
}
