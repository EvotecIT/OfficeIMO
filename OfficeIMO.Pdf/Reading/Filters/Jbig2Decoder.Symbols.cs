namespace OfficeIMO.Pdf.Filters;

internal static partial class Jbig2Decoder {
    /// <summary>Value source for the text region decoding procedure (6.4): arithmetic contexts or Huffman tables.</summary>
    private sealed class TextCoding {
        internal ArithmeticDecoder? Arithmetic;
        internal IntegerContexts? Integers;
        internal byte[]? IdContexts;
        internal HuffmanReader? Reader;
        internal HuffmanTable? FirstS, DeltaS, DeltaT, RefinementWidth, RefinementHeight, RefinementX, RefinementY, RefinementSize;
        internal HuffmanTable? SymbolIds; // null with Huffman coding means fixed-length IDs (6.5.8.2.3)
        internal int CodeLength;
        internal byte[] RefinementContexts = Array.Empty<byte>();
        internal byte[] Data = Array.Empty<byte>();

        internal bool Huffman => Reader != null;

        internal int? Integer(string arithmeticName, HuffmanTable? table) =>
            Huffman ? table!.Decode(Reader!) : Arithmetic!.DecodeInteger(Integers![arithmeticName]);

        internal int Required(string arithmeticName, HuffmanTable? table, string description) =>
            Integer(arithmeticName, table) ?? throw new InvalidDataException("JBIG2 " + description + " is out of band.");

        internal int Id() {
            if (!Huffman) return Arithmetic!.DecodeId(IdContexts!, CodeLength);
            return SymbolIds != null ? SymbolIds.DecodeValue(Reader!, "symbol ID") : (int)Reader!.ReadBits(CodeLength);
        }
    }

    private sealed partial class DecodingContext {
        private readonly Dictionary<int, HuffmanTable> _tables = new();

        private void ProcessTable(Segment segment) => _tables[segment.Number] = ReadCustomTable(segment, Charge);

        private List<HuffmanTable> ReferredTables(Segment segment) {
            var tables = new List<HuffmanTable>();
            foreach (int referred in segment.Referred)
                if (_tables.TryGetValue(referred, out HuffmanTable? table)) tables.Add(table);
            return tables;
        }

        // Symbol dictionary segment (7.4.2) and decoding procedure (6.5.5).
        private void ProcessSymbolDictionary(Segment segment) {
            int position = segment.DataStart;
            int flags = (ReadByte(segment.Data, ref position, segment.DataEnd) << 8) | ReadByte(segment.Data, ref position, segment.DataEnd);
            bool huffman = (flags & 1) != 0;
            bool refinementAggregate = (flags & 2) != 0;
            int template = (flags >> 10) & 3;
            int refinementTemplate = (flags >> 12) & 1;
            sbyte[] at = huffman ? Array.Empty<sbyte>() : ReadAt(segment.Data, ref position, template == 0 ? 4 : 1, segment.DataEnd);
            sbyte[] refinementAt = refinementAggregate && refinementTemplate == 0 ? ReadAt(segment.Data, ref position, 2, segment.DataEnd) : Array.Empty<sbyte>();
            int exportedCount = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int newCount = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            List<Bitmap> input = ReferredSymbols(segment);
            if (newCount < 0 || exportedCount < 0 || (long)newCount + input.Count > MaximumSymbols) throw new InvalidDataException("JBIG2 symbol dictionary is too large.");

            ArithmeticDecoder? decoder = null;
            HuffmanReader? reader = null;
            HuffmanTable? heightTable = null, widthTable = null, sizeTable = null, aggregateTable = null;
            if (huffman) {
                var cursor = new CustomTableCursor(ReferredTables(segment));
                heightTable = cursor.Select((flags >> 2) & 3, 3, 4, 5);
                widthTable = cursor.Select((flags >> 4) & 3, 3, 2, 3);
                sizeTable = cursor.Select((flags >> 6) & 1, 1, 1);
                aggregateTable = cursor.Select((flags >> 7) & 1, 1, 1);
                reader = new HuffmanReader(segment.Data, position, segment.DataEnd, Charge);
            } else {
                decoder = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
            }
            var integers = new IntegerContexts(CreateContexts);
            SymbolBitmapContexts bitmapContexts = CreateSymbolBitmapContexts(segment, flags, at, refinementAt);
            byte[] genericContexts = bitmapContexts.Generic;
            byte[] refinementContexts = bitmapContexts.Refinement;
            int codeLength = CodeLength(input.Count + newCount);
            byte[] idContexts = huffman ? Array.Empty<byte>() : CreateContexts(1 << (codeLength + 1));
            var symbols = CreateSymbolList(input.Count + newCount);
            symbols.AddRange(input);
            int heightClass = 0;
            int decoded = 0;
            var classWidths = new List<int>();
            while (decoded < newCount) {
                _cancellationToken.ThrowIfCancellationRequested();
                int deltaHeight = huffman ? heightTable!.DecodeValue(reader!, "height class delta")
                    : decoder!.DecodeInteger(integers["IADH"]) ?? throw new InvalidDataException("JBIG2 height class delta is out of band.");
                heightClass += deltaHeight;
                if (heightClass < 0) throw new InvalidDataException("JBIG2 symbol height is negative.");
                int symbolWidth = 0;
                long totalWidth = 0;
                classWidths.Clear();
                while (true) {
                    int? deltaWidth = huffman ? widthTable!.Decode(reader!) : decoder!.DecodeInteger(integers["IADW"]);
                    if (deltaWidth == null) break;
                    symbolWidth += deltaWidth.Value;
                    if (symbolWidth < 0 || decoded >= newCount) throw new InvalidDataException("JBIG2 symbol width is invalid.");
                    if (huffman && !refinementAggregate) {
                        // Collected into one height-class bitmap after the class ends (6.5.9).
                        classWidths.Add(symbolWidth);
                        totalWidth += symbolWidth;
                        decoded++;
                        continue;
                    }
                    Bitmap symbol;
                    if (!refinementAggregate) {
                        Charge((long)symbolWidth * heightClass);
                        symbol = DecodeGeneric(decoder!, genericContexts, symbolWidth, heightClass, template, false, at);
                    } else {
                        int instances = huffman ? aggregateTable!.DecodeValue(reader!, "aggregate count")
                            : decoder!.DecodeInteger(integers["IAAI"]) ?? throw new InvalidDataException("JBIG2 aggregate count is out of band.");
                        if (instances == 1) {
                            int id, dx, dy;
                            if (huffman) {
                                id = (int)reader!.ReadBits(codeLength);
                                dx = Standard(15).DecodeValue(reader, "refinement X");
                                dy = Standard(15).DecodeValue(reader, "refinement Y");
                                int size = Standard(1).DecodeValue(reader, "refinement size");
                                reader.Align();
                                if ((uint)id >= (uint)symbols.Count) throw new InvalidDataException("JBIG2 refinement refers to an unknown symbol.");
                                int start = reader.BytePosition;
                                if (size <= 0 || size > segment.DataEnd - start) throw new InvalidDataException("JBIG2 refinement size is invalid.");
                                var refinementDecoder = new ArithmeticDecoder(segment.Data, start, start + size, Charge);
                                symbol = DecodeRefinement(refinementDecoder, refinementContexts, symbolWidth, heightClass, refinementTemplate, symbols[id], dx, dy, false, refinementAt);
                                reader.SkipBytes(size);
                            } else {
                                id = decoder!.DecodeId(idContexts, codeLength);
                                dx = decoder.DecodeInteger(integers["IARDX"]) ?? 0;
                                dy = decoder.DecodeInteger(integers["IARDY"]) ?? 0;
                                if ((uint)id >= (uint)symbols.Count) throw new InvalidDataException("JBIG2 refinement refers to an unknown symbol.");
                                symbol = DecodeRefinement(decoder, refinementContexts, symbolWidth, heightClass, refinementTemplate, symbols[id], dx, dy, false, refinementAt);
                            }
                        } else if (instances > 1) {
                            TextCoding coding = huffman
                                ? new TextCoding {
                                    Reader = reader, FirstS = Standard(6), DeltaS = Standard(8), DeltaT = Standard(11),
                                    RefinementWidth = Standard(15), RefinementHeight = Standard(15), RefinementX = Standard(15),
                                    RefinementY = Standard(15), RefinementSize = Standard(1), CodeLength = codeLength
                                }
                                : new TextCoding { Arithmetic = decoder, Integers = integers, IdContexts = idContexts, CodeLength = codeLength };
                            coding.RefinementContexts = refinementContexts;
                            coding.Data = segment.Data;
                            symbol = DecodeTextRegion(coding, symbolWidth, heightClass, instances, symbols, stripLog: 0, referenceCorner: 1,
                                transposed: false, combination: 0, defaultPixel: 0, dsOffset: 0, refine: true, refinementTemplate, refinementAt);
                        } else {
                            throw new InvalidDataException("JBIG2 aggregate count is invalid.");
                        }
                    }
                    symbols.Add(symbol);
                    decoded++;
                }
                if (huffman && !refinementAggregate) AddCollectiveSymbols(segment, reader!, sizeTable!, classWidths, totalWidth, heightClass, symbols);
            }
            // Export flags (6.5.10) select from input then new symbols.
            if (exportedCount > symbols.Count) throw new InvalidDataException("JBIG2 exported symbol count is invalid.");
            var exported = CreateSymbolList(exportedCount);
            bool exporting = false;
            int index = 0;
            while (index < symbols.Count) {
                int run = huffman ? Standard(1).DecodeValue(reader!, "export run")
                    : decoder!.DecodeInteger(integers["IAEX"]) ?? throw new InvalidDataException("JBIG2 export run is out of band.");
                if (run < 0 || run > symbols.Count - index) throw new InvalidDataException("JBIG2 export run is invalid.");
                if (exporting) {
                    if (run > exportedCount - exported.Count) throw new InvalidDataException("JBIG2 export runs exceed the declared count.");
                    for (int exportedIndex = index; exportedIndex < index + run; exportedIndex++) {
                        Charge(1L);
                        exported.Add(symbols[exportedIndex]);
                    }
                }
                index += run;
                exporting = !exporting;
            }
            if (exported.Count != exportedCount) throw new InvalidDataException("JBIG2 export runs do not match the declared count.");
            _symbols[segment.Number] = exported;
            if ((flags & 0x200) != 0) _retainedSymbolContexts[segment.Number] = bitmapContexts;
        }

        // Height class collective bitmap (6.5.9): uncompressed when BMSIZE is zero, otherwise MMR.
        private void AddCollectiveSymbols(Segment segment, HuffmanReader reader, HuffmanTable sizeTable, List<int> widths,
            long totalWidth, int height, List<Bitmap> symbols) {
            int size = sizeTable.DecodeValue(reader, "collective bitmap size");
            reader.Align();
            if (totalWidth > MaximumBitmapPixels || totalWidth * height > MaximumBitmapPixels) throw new InvalidDataException("JBIG2 height class is too large.");
            int width = (int)totalWidth;
            Bitmap collective;
            int start = reader.BytePosition;
            if (size == 0) {
                long stride = ((long)width + 7) / 8;
                if (stride * height > segment.DataEnd - start) throw new InvalidDataException("Truncated JBIG2 collective bitmap.");
                collective = CreateBitmap(width, height);
                for (int y = 0; y < height; y++) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    for (int x = 0; x < width; x++)
                        if (((segment.Data[start + y * stride + x / 8] >> (7 - (x & 7))) & 1) != 0) collective.Set(x, y, 1);
                }
                reader.SkipBytes((int)(stride * height));
            } else {
                if (size < 0 || size > segment.DataEnd - start) throw new InvalidDataException("JBIG2 collective bitmap size is invalid.");
                ReserveBytes(size);
                var encoded = new byte[size];
                Buffer.BlockCopy(segment.Data, start, encoded, 0, size);
                collective = DecodeMmr(encoded, width, height);
                reader.SkipBytes(size);
            }
            int left = 0;
            foreach (int symbolWidth in widths) {
                Bitmap symbol = CreateBitmap(symbolWidth, height);
                symbol.CopyFrom(collective, left, 0);
                symbols.Add(symbol);
                left += symbolWidth;
            }
        }

        // Text region segment (7.4.3).
        private void ProcessTextRegion(Segment segment) {
            int position = segment.DataStart;
            var info = ReadRegionInfo(segment, ref position);
            int flags = (ReadByte(segment.Data, ref position, segment.DataEnd) << 8) | ReadByte(segment.Data, ref position, segment.DataEnd);
            bool huffman = (flags & 1) != 0;
            bool refine = (flags & 2) != 0;
            int stripLog = (flags >> 2) & 3;
            int referenceCorner = (flags >> 4) & 3;
            bool transposed = (flags & 0x40) != 0;
            int combination = (flags >> 7) & 3;
            int defaultPixel = (flags >> 9) & 1;
            int dsOffset = (flags >> 10) & 0x1F;
            if ((dsOffset & 0x10) != 0) dsOffset -= 32;
            int refinementTemplate = (flags >> 15) & 1;
            int huffmanFlags = huffman
                ? (ReadByte(segment.Data, ref position, segment.DataEnd) << 8) | ReadByte(segment.Data, ref position, segment.DataEnd)
                : 0;
            sbyte[] refinementAt = refine && refinementTemplate == 0 ? ReadAt(segment.Data, ref position, 2, segment.DataEnd) : Array.Empty<sbyte>();
            int instances = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            if (instances < 0 || instances > MaximumSymbols * 4) throw new InvalidDataException("JBIG2 text region has too many instances.");
            List<Bitmap> symbols = ReferredSymbols(segment);
            int codeLength = CodeLength(symbols.Count);
            var coding = new TextCoding { CodeLength = codeLength, RefinementContexts = CreateContexts(1 << 14), Data = segment.Data };
            if (huffman) {
                var cursor = new CustomTableCursor(ReferredTables(segment));
                coding.FirstS = cursor.Select(huffmanFlags & 3, 3, 6, 7);
                coding.DeltaS = cursor.Select((huffmanFlags >> 2) & 3, 3, 8, 9, 10);
                coding.DeltaT = cursor.Select((huffmanFlags >> 4) & 3, 3, 11, 12, 13);
                coding.RefinementWidth = cursor.Select((huffmanFlags >> 6) & 3, 3, 14, 15);
                coding.RefinementHeight = cursor.Select((huffmanFlags >> 8) & 3, 3, 14, 15);
                coding.RefinementX = cursor.Select((huffmanFlags >> 10) & 3, 3, 14, 15);
                coding.RefinementY = cursor.Select((huffmanFlags >> 12) & 3, 3, 14, 15);
                coding.RefinementSize = cursor.Select((huffmanFlags >> 14) & 1, 1, 1);
                coding.Reader = new HuffmanReader(segment.Data, position, segment.DataEnd, Charge);
                coding.SymbolIds = ReadSymbolIdTable(coding.Reader, symbols.Count);
            } else {
                coding.Arithmetic = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
                coding.Integers = new IntegerContexts(CreateContexts);
                coding.IdContexts = CreateContexts(1 << (codeLength + 1));
            }
            Charge((long)info.Width * info.Height);
            Bitmap region = DecodeTextRegion(coding, info.Width, info.Height, instances, symbols, stripLog, referenceCorner, transposed,
                combination, defaultPixel, dsOffset, refine, refinementTemplate, refinementAt);
            Place(segment, region, info);
        }

        // Symbol ID Huffman decoding table (7.4.3.1.7).
        private HuffmanTable ReadSymbolIdTable(HuffmanReader reader, int symbolCount) {
            var runLines = new HuffmanLine[35];
            for (int index = 0; index < runLines.Length; index++) runLines[index] = L((int)reader.ReadBits(4), 0, index);
            var runCodes = new HuffmanTable(runLines);
            ReserveBytes((long)symbolCount * 16);
            var lengths = new int[symbolCount];
            int position = 0;
            while (position < symbolCount) {
                _cancellationToken.ThrowIfCancellationRequested();
                int code = runCodes.DecodeValue(reader, "symbol ID run code");
                int repeat, value;
                if (code < 32) {
                    lengths[position++] = code;
                    continue;
                }
                if (code == 32) {
                    if (position == 0) throw new InvalidDataException("JBIG2 symbol ID run repeats a missing length.");
                    value = lengths[position - 1];
                    repeat = 3 + (int)reader.ReadBits(2);
                } else if (code == 33) {
                    value = 0;
                    repeat = 3 + (int)reader.ReadBits(3);
                } else {
                    value = 0;
                    repeat = 11 + (int)reader.ReadBits(7);
                }
                if (repeat > symbolCount - position) throw new InvalidDataException("JBIG2 symbol ID run exceeds the symbol count.");
                for (int index = 0; index < repeat; index++) lengths[position++] = value;
            }
            reader.Align();
            var lines = new HuffmanLine[symbolCount];
            for (int index = 0; index < symbolCount; index++) lines[index] = L(lengths[index], 0, index);
            return new HuffmanTable(lines);
        }

        // Text region decoding procedure (6.4.5). Reference corners: 0 bottom-left, 1 top-left, 2 bottom-right, 3 top-right.
        private Bitmap DecodeTextRegion(TextCoding coding, int width, int height, int instances, List<Bitmap> symbols, int stripLog,
            int referenceCorner, bool transposed, int combination, int defaultPixel, int dsOffset, bool refine, int refinementTemplate,
            sbyte[] refinementAt) {
            var region = CreateBitmap(width, height);
            if (defaultPixel != 0) region.Fill(1);
            int strips = 1 << stripLog;
            long stripT = -(long)coding.Required("IADT", coding.DeltaT, "strip delta") * strips;
            long firstS = 0;
            int count = 0;
            while (count < instances) {
                _cancellationToken.ThrowIfCancellationRequested();
                stripT += (long)coding.Required("IADT", coding.DeltaT, "strip delta") * strips;
                firstS += coding.Required("IAFS", coding.FirstS, "first S");
                long currentS = firstS;
                while (true) {
                    int currentT = strips == 1 ? 0 : coding.Huffman ? (int)coding.Reader!.ReadBits(stripLog) : coding.Integer("IAIT", null) ?? 0;
                    long t = stripT + currentT;
                    int id = coding.Id();
                    if ((uint)id >= (uint)symbols.Count) throw new InvalidDataException("JBIG2 text region refers to an unknown symbol.");
                    Bitmap symbol = symbols[id];
                    int refinement = !refine ? 0 : coding.Huffman ? coding.Reader!.ReadBit() : coding.Integer("IARI", null) ?? 0;
                    if (refinement != 0) symbol = RefineSymbol(coding, symbol, refinementTemplate, refinementAt);
                    int symbolWidth = symbol.Width, symbolHeight = symbol.Height;
                    if (!transposed && referenceCorner >= 2) currentS += symbolWidth - 1;
                    else if (transposed && (referenceCorner & 1) == 0) currentS += symbolHeight - 1;
                    long s = currentS;
                    long x, y;
                    if (!transposed) {
                        x = referenceCorner >= 2 ? s - symbolWidth + 1 : s;
                        y = (referenceCorner & 1) == 1 ? t : t - symbolHeight + 1;
                    } else {
                        x = referenceCorner >= 2 ? t - symbolWidth + 1 : t;
                        y = (referenceCorner & 1) == 1 ? s : s - symbolHeight + 1;
                    }
                    if (x > int.MinValue && x < int.MaxValue && y > int.MinValue && y < int.MaxValue)
                        region.Compose(symbol, (int)x, (int)y, combination);
                    if (!transposed && referenceCorner < 2) currentS += symbolWidth - 1;
                    else if (transposed && (referenceCorner & 1) == 1) currentS += symbolHeight - 1;
                    count++;
                    // Every strip ends with an out-of-band IADS, also after the final instance (6.4.5 step 3c). Reading
                    // it keeps a shared decoder aligned when an aggregate symbol is decoded inside a dictionary.
                    int? deltaS = ReadFinalDeltaS(coding, count >= instances);
                    if (deltaS == null || count >= instances) break;
                    currentS += deltaS.Value + dsOffset;
                }
            }
            return region;
        }

        private static int? ReadFinalDeltaS(TextCoding coding, bool last) {
            if (!last || !coding.Huffman) return coding.Integer("IADS", coding.DeltaS);
            // A Huffman text region may end without the final out-of-band code; nothing follows it in the segment.
            try { return coding.Integer("IADS", coding.DeltaS); }
            catch (InvalidDataException) { return null; }
        }

        // Refined symbol instance (6.4.11.3).
        private Bitmap RefineSymbol(TextCoding coding, Bitmap symbol, int refinementTemplate, sbyte[] refinementAt) {
            int deltaWidth = coding.Integer("IARDW", coding.RefinementWidth) ?? 0;
            int deltaHeight = coding.Integer("IARDH", coding.RefinementHeight) ?? 0;
            int dx = coding.Integer("IARDX", coding.RefinementX) ?? 0;
            int dy = coding.Integer("IARDY", coding.RefinementY) ?? 0;
            long refinedWidth = (long)symbol.Width + deltaWidth, refinedHeight = (long)symbol.Height + deltaHeight;
            if (refinedWidth < 0 || refinedHeight < 0 || refinedWidth * refinedHeight > MaximumBitmapPixels)
                throw new InvalidDataException("JBIG2 refined symbol size is invalid.");
            int referenceDx = (deltaWidth >> 1) + dx, referenceDy = (deltaHeight >> 1) + dy;
            if (!coding.Huffman) {
                return DecodeRefinement(coding.Arithmetic!, coding.RefinementContexts, (int)refinedWidth, (int)refinedHeight,
                    refinementTemplate, symbol, referenceDx, referenceDy, false, refinementAt);
            }
            HuffmanReader reader = coding.Reader!;
            int size = coding.RefinementSize!.DecodeValue(reader, "refinement size");
            reader.Align();
            int start = reader.BytePosition;
            if (size <= 0 || (long)start + size > reader.End) throw new InvalidDataException("JBIG2 refinement size is invalid.");
            var decoder = new ArithmeticDecoder(coding.Data, start, start + size, Charge);
            Bitmap refined = DecodeRefinement(decoder, coding.RefinementContexts, (int)refinedWidth, (int)refinedHeight,
                refinementTemplate, symbol, referenceDx, referenceDy, false, refinementAt);
            reader.SkipBytes(size);
            return refined;
        }
    }
}
