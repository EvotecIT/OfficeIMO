namespace OfficeIMO.DjVu;

internal static class DjVuTextReader {
    internal static DjVuTextResult Read(DjVuChunk? chunk, DjVuReadBudget budget) {
        if (chunk == null) return new DjVuTextResult(DjVuTextStatus.Absent, string.Empty);
        try {
            byte[] data;
            int start, end;
            if (chunk.Id == "TXTz") {
                data = BzzDecoder.Decode(chunk.Source, chunk.Offset, chunk.Length, budget);
                start = 0; end = data.Length;
            } else {
                budget.Expanded(chunk.Length);
                data = chunk.Source; start = chunk.Offset; end = start + chunk.Length;
            }
            if (end - start < 3) throw new InvalidDataException("Truncated DjVu text header.");
            int textLength = DjVuBinary.U24(data, start), textStart = start + 3;
            if (textLength > end - textStart) throw new InvalidDataException("Stored DjVu text exceeds its chunk.");
            string text = DjVuBinary.Utf8.GetString(data, textStart, textLength);
            budget.TextCharacters(text.Length);
            var zones = new List<DjVuTextZone>();
            int position = textStart + textLength;
            if (position < end) {
                if (data[position++] != 1) throw new InvalidDataException("Unsupported DjVu text-zone version.");
                if (position < end) {
                    budget.WorkingBytes((textLength + 1L) * 4 + text.Length * 2L + end - start);
                    int[] offsets = CharacterOffsets(data, textStart, textLength, budget.Cancellation);
                    DjVuTextZone? previous = null;
                    while (position < end) {
                        var zone = ReadZone(data, ref position, end, offsets, null, previous, budget, 1);
                        zones.Add(zone); previous = zone;
                    }
                }
            }
            return new DjVuTextResult(text.Length == 0 ? DjVuTextStatus.Empty : DjVuTextStatus.Present, text, zones);
        } catch (InvalidDataException exception) {
            return new DjVuTextResult(DjVuTextStatus.Corrupt, string.Empty, diagnostic: exception.Message);
        } catch (DecoderFallbackException exception) {
            return new DjVuTextResult(DjVuTextStatus.Corrupt, string.Empty, diagnostic: exception.Message);
        } catch (OverflowException) {
            return new DjVuTextResult(DjVuTextStatus.Corrupt, string.Empty, diagnostic: "DjVu text-zone coordinates overflow.");
        }
    }

    private static int[] CharacterOffsets(byte[] data, int start, int length, CancellationToken cancellation) {
        var offsets = new int[length + 1];
        int characters = 0;
        for (int i = 0; i < length;) {
            if ((i & 4095) == 0) cancellation.ThrowIfCancellationRequested();
            byte first = data[start + i];
            int width = first < 128 ? 1 : first < 224 ? 2 : first < 240 ? 3 : 4;
            offsets[i] = characters;
            for (int j = 1; j < width; j++) offsets[i + j] = -1;
            i += width;
            characters += width == 4 ? 2 : 1;
        }
        offsets[length] = characters;
        return offsets;
    }

    private static DjVuTextZone ReadZone(byte[] data, ref int position, int end, int[] characters,
        DjVuTextZone? parent, DjVuTextZone? previous, DjVuReadBudget budget, int depth) {
        budget.TextZone(depth);
        if (end - position < 17) throw new InvalidDataException("Truncated DjVu text zone.");
        int kind = data[position++];
        if (kind < 1 || kind > 7 || parent != null && kind <= (int)parent.Kind) throw new InvalidDataException("Invalid DjVu text-zone hierarchy.");
        int x = DjVuBinary.U16(data, position) - 32768;
        int y = DjVuBinary.U16(data, position + 2) - 32768;
        int width = DjVuBinary.U16(data, position + 4) - 32768;
        int height = DjVuBinary.U16(data, position + 6) - 32768;
        int relativeText = DjVuBinary.U16(data, position + 8) - 32768;
        int length = DjVuBinary.U24(data, position + 10);
        int childCount = DjVuBinary.U24(data, position + 13);
        position += 16;
        if (width < 0 || height < 0 || childCount > (end - position) / 17) throw new InvalidDataException("Invalid DjVu text-zone extent or child count.");
        int textOffset = relativeText;
        checked {
            if (previous != null) {
                textOffset += previous.ByteOffset + previous.ByteLength;
                if (kind == 1 || kind == 4 || kind == 5) {
                    x += previous.Bounds.X;
                    y = previous.Bounds.Y - y - height;
                } else {
                    x += previous.Bounds.X + previous.Bounds.Width;
                    y += previous.Bounds.Y;
                }
            } else if (parent != null) {
                textOffset += parent.ByteOffset;
                x += parent.Bounds.X;
                y = parent.Bounds.Y + parent.Bounds.Height - y - height;
            }
        }
        if (textOffset < 0 || textOffset > characters.Length - 1 - length || characters[textOffset] < 0 || characters[textOffset + length] < 0)
            throw new InvalidDataException("DjVu text zone is outside text or splits a UTF-8 character.");
        if (parent != null && (textOffset < parent.ByteOffset || length > parent.ByteOffset + parent.ByteLength - textOffset))
            throw new InvalidDataException("DjVu child text is outside its parent zone.");
        var children = new List<DjVuTextZone>();
        var zone = new DjVuTextZone((DjVuTextZoneKind)kind, new DjVuRectangle(x, y, width, height), textOffset, length,
            characters[textOffset], characters[textOffset + length] - characters[textOffset], children);
        DjVuTextZone? sibling = null;
        for (int i = 0; i < childCount; i++) {
            var child = ReadZone(data, ref position, end, characters, zone, sibling, budget, depth + 1);
            children.Add(child); sibling = child;
        }
        return zone;
    }
}
