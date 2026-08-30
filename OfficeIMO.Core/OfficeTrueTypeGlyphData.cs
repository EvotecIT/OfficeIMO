using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared bounded structural validation for TrueType <c>glyf</c> records.</summary>
internal static class OfficeTrueTypeGlyphData {
    private const ushort ArgumentsAreWords = 0x0001;
    private const ushort MoreComponents = 0x0020;
    private const ushort HasScale = 0x0008;
    private const ushort HasXAndYScale = 0x0040;
    private const ushort HasTwoByTwo = 0x0080;
    private const ushort HasInstructions = 0x0100;
    private const ushort UseMyMetrics = 0x0200;
    private const ushort SupportedCompositeFlags = 0x1FFF;

    internal static bool IsStructurallyValid(
        byte[] data,
        int glyfOffset,
        int glyfLength,
        IReadOnlyList<uint> locations) {
        if (data == null || locations == null || locations.Count < 2 ||
            glyfOffset < 0 || glyfLength <= 0 || glyfOffset > data.Length - glyfLength) return false;

        int glyphCount = locations.Count - 1;
        var componentReferences = new int[glyphCount][];
        for (int glyph = 0; glyph < glyphCount; glyph++) {
            uint start = locations[glyph];
            uint end = locations[glyph + 1];
            if (start > end || end > glyfLength) return false;
            if (start == end) {
                componentReferences[glyph] = Array.Empty<int>();
                continue;
            }

            int glyphLength = checked((int)(end - start));
            int glyphOffset = checked(glyfOffset + (int)start);
            if (glyphLength < 10 || glyphOffset < glyfOffset || glyphOffset > data.Length - glyphLength) return false;
            if (ReadInt16(data, glyphOffset + 2) > ReadInt16(data, glyphOffset + 6) ||
                ReadInt16(data, glyphOffset + 4) > ReadInt16(data, glyphOffset + 8)) return false;
            short contourCount = ReadInt16(data, glyphOffset);
            if (contourCount >= 0) {
                if (!IsValidSimpleGlyph(data, glyphOffset, glyphLength, contourCount)) return false;
                componentReferences[glyph] = Array.Empty<int>();
            } else {
                if (!TryReadCompositeGlyph(data, glyphOffset, glyphLength, glyphCount, out int[] references)) return false;
                componentReferences[glyph] = references;
            }
        }

        return HasAcyclicCompositeReferences(componentReferences);
    }

    private static bool IsValidSimpleGlyph(byte[] data, int glyphOffset, int glyphLength, int contourCount) {
        int glyphEnd = checked(glyphOffset + glyphLength);
        long endPointsEnd = (long)glyphOffset + 10L + contourCount * 2L;
        if (endPointsEnd > glyphEnd - 2L) return false;

        int cursor = glyphOffset + 10;
        int pointCount = 0;
        int previousEndPoint = -1;
        for (int contour = 0; contour < contourCount; contour++) {
            int endPoint = ReadUInt16(data, cursor);
            cursor += 2;
            if (endPoint <= previousEndPoint) return false;
            previousEndPoint = endPoint;
        }
        if (contourCount > 0) pointCount = checked(previousEndPoint + 1);

        int instructionLength = ReadUInt16(data, cursor);
        cursor += 2;
        if (instructionLength > glyphEnd - cursor) return false;
        cursor += instructionLength;

        var flags = new byte[pointCount];
        int flagIndex = 0;
        while (flagIndex < pointCount) {
            if (cursor >= glyphEnd) return false;
            byte flag = data[cursor++];
            if ((flag & 0xC0) != 0) return false;
            flags[flagIndex++] = flag;
            if ((flag & 0x08) == 0) continue;
            if (cursor >= glyphEnd) return false;
            int repeatCount = data[cursor++];
            if (repeatCount > pointCount - flagIndex) return false;
            for (int repeat = 0; repeat < repeatCount; repeat++) flags[flagIndex++] = flag;
        }

        long coordinateBytes = 0L;
        foreach (byte flag in flags) {
            coordinateBytes += (flag & 0x02) != 0 ? 1 : (flag & 0x10) == 0 ? 2 : 0;
            coordinateBytes += (flag & 0x04) != 0 ? 1 : (flag & 0x20) == 0 ? 2 : 0;
        }
        return coordinateBytes <= glyphEnd - cursor;
    }

    private static bool TryReadCompositeGlyph(
        byte[] data,
        int glyphOffset,
        int glyphLength,
        int glyphCount,
        out int[] references) {
        references = Array.Empty<int>();
        int glyphEnd = checked(glyphOffset + glyphLength);
        int cursor = glyphOffset + 10;
        var components = new List<int>();
        ushort flags;
        bool hasMetricsComponent = false;
        do {
            if (cursor > glyphEnd - 4) return false;
            flags = ReadUInt16(data, cursor);
            int component = ReadUInt16(data, cursor + 2);
            cursor += 4;
            if ((flags & ~SupportedCompositeFlags) != 0 ||
                OfficeOpenTypeCompositeGlyph.HasConflictingTransformFlags(flags) ||
                OfficeOpenTypeCompositeGlyph.HasConflictingOffsetFlags(flags) ||
                (flags & MoreComponents) != 0 && (flags & HasInstructions) != 0 ||
                component >= glyphCount) return false;

            if ((flags & UseMyMetrics) != 0) {
                if (hasMetricsComponent) return false;
                hasMetricsComponent = true;
            }
            components.Add(component);

            int argumentBytes = (flags & ArgumentsAreWords) != 0 ? 4 : 2;
            int transformBytes = (flags & HasScale) != 0
                ? 2
                : (flags & HasXAndYScale) != 0
                    ? 4
                    : (flags & HasTwoByTwo) != 0 ? 8 : 0;
            int componentBytes = checked(argumentBytes + transformBytes);
            if (componentBytes > glyphEnd - cursor) return false;
            cursor += componentBytes;
        } while ((flags & MoreComponents) != 0);

        if ((flags & HasInstructions) != 0) {
            if (cursor > glyphEnd - 2) return false;
            int instructionLength = ReadUInt16(data, cursor);
            cursor += 2;
            if (instructionLength > glyphEnd - cursor) return false;
        }

        references = components.ToArray();
        return references.Length > 0;
    }

    private static bool HasAcyclicCompositeReferences(IReadOnlyList<int[]> references) {
        var states = new byte[references.Count];
        var nextEdges = new int[references.Count];
        var stack = new Stack<int>();
        for (int start = 0; start < references.Count; start++) {
            if (states[start] != 0 || references[start].Length == 0) continue;
            states[start] = 1;
            stack.Push(start);
            while (stack.Count > 0) {
                int glyph = stack.Peek();
                int edge = nextEdges[glyph];
                if (edge >= references[glyph].Length) {
                    states[glyph] = 2;
                    stack.Pop();
                    continue;
                }

                int component = references[glyph][edge];
                nextEdges[glyph] = edge + 1;
                if (states[component] == 1) return false;
                if (states[component] != 0 || references[component].Length == 0) continue;
                states[component] = 1;
                stack.Push(component);
            }
        }
        return true;
    }

    private static short ReadInt16(byte[] data, int offset) => unchecked((short)ReadUInt16(data, offset));

    private static ushort ReadUInt16(byte[] data, int offset) =>
        (ushort)((data[offset] << 8) | data[offset + 1]);
}
