namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    private const int MaximumCodeBlockPasses = 164;

    /// <summary>Packet data of one tile: bodies and, with PPM or PPT, separately stored packet headers.</summary>
    private sealed class TileStream {
        internal byte[] Data = Array.Empty<byte>();
        internal int Position;
        internal int End;
        internal PacketHeaderReader? PackedHeaders;
        internal PacketHeaderReader? BodyHeaders;
        internal bool Sop, Eph;
        internal bool Exhausted;
        internal readonly List<(CodeSegment Segment, int Length)> Pending = new();
    }

    /// <summary>Packet-header bit reader with bit-stuffing after 0xFF (B.10.1).</summary>
    private sealed class PacketHeaderReader {
        private readonly byte[] _data;
        private readonly int _end;
        private int _current;
        private int _bitsLeft;

        internal PacketHeaderReader(byte[] data, int start, int end) {
            _data = data;
            Position = start;
            _end = end;
        }

        internal int Position { get; set; }

        internal void StartPacket() {
            _current = 0;
            _bitsLeft = 0;
        }

        internal int ReadBit() {
            if (_bitsLeft == 0) {
                if (Position >= _end) throw new PacketDataExhaustedException();
                _bitsLeft = _current == 0xFF ? 7 : 8;
                _current = _data[Position++];
            }
            _bitsLeft--;
            return (_current >> _bitsLeft) & 1;
        }

        internal int ReadBits(int count) {
            int value = 0;
            for (int index = 0; index < count; index++) value = (value << 1) | ReadBit();
            return value;
        }

        // A header ending in 0xFF is followed by one stuffed byte that still belongs to the header.
        internal void Align() {
            if (_current == 0xFF) {
                if (Position >= _end) throw new PacketDataExhaustedException();
                Position++;
            }
            _current = 0;
            _bitsLeft = 0;
        }

        internal void SkipMarker(int marker) {
            if (_end - Position >= 2 && ((_data[Position] << 8) | _data[Position + 1]) == marker) Position += 2;
        }
    }

    private sealed class PacketDataExhaustedException : Exception { }

    private static TileStream CreateTileStream(byte[] bytes, Codestream codestream, TileData tile, GlobalCoding coding, Budget budget) {
        var stream = new TileStream { Sop = coding.Sop, Eph = coding.Eph };
        if (tile.Bodies.Count == 1) {
            stream.Data = bytes;
            stream.Position = tile.Bodies[0].Start;
            stream.End = tile.Bodies[0].End;
        } else {
            long total = 0;
            foreach (var body in tile.Bodies) total += body.End - body.Start;
            budget.Reserve(total);
            stream.Data = new byte[total];
            int offset = 0;
            foreach (var body in tile.Bodies) {
                Buffer.BlockCopy(bytes, body.Start, stream.Data, offset, body.End - body.Start);
                offset += body.End - body.Start;
            }
            stream.End = offset;
        }
        byte[]? headers = null;
        if (tile.PackedHeaders.Count > 0) {
            long total = 0;
            foreach (var record in tile.PackedHeaders) total += record.End - record.Start;
            budget.Reserve(total);
            headers = new byte[total];
            int offset = 0;
            foreach (var record in tile.PackedHeaders) {
                Buffer.BlockCopy(codestream.PackedHeaderData!, record.Start, headers, offset, record.End - record.Start);
                offset += record.End - record.Start;
            }
        } else if (tile.PacketHeaderChunks.Count > 0) {
            tile.PacketHeaderChunks.Sort(static (left, right) => left.Index.CompareTo(right.Index));
            long total = 0;
            foreach (var chunk in tile.PacketHeaderChunks) total += chunk.Data.Length;
            budget.Reserve(total);
            headers = new byte[total];
            int offset = 0;
            foreach (var chunk in tile.PacketHeaderChunks) {
                Buffer.BlockCopy(chunk.Data, 0, headers, offset, chunk.Data.Length);
                offset += chunk.Data.Length;
            }
        }
        if (headers != null) stream.PackedHeaders = new PacketHeaderReader(headers, 0, headers.Length);
        else stream.BodyHeaders = new PacketHeaderReader(stream.Data, stream.Position, stream.End);
        return stream;
    }

    // Progression orders (B.12) with POC volumes, followed by the default order for any packet a volume omitted.
    private static void DecodePackets(TileStream stream, TileComponent[] components, GlobalCoding coding,
        List<ProgressionChange> changes, Budget budget) {
        int maximumResolutions = 0;
        foreach (TileComponent component in components) maximumResolutions = Math.Max(maximumResolutions, component.Coding.Levels + 1);
        var volumes = new List<ProgressionChange>(changes.Count + 1);
        volumes.AddRange(changes);
        volumes.Add(new ProgressionChange {
            ResolutionStart = 0, ResolutionEnd = maximumResolutions, ComponentStart = 0, ComponentEnd = components.Length,
            LayerEnd = coding.Layers, Order = coding.Progression
        });
        try {
            foreach (ProgressionChange volume in volumes) {
                int layerEnd = Math.Min(volume.LayerEnd, coding.Layers);
                int resolutionEnd = Math.Min(volume.ResolutionEnd, maximumResolutions);
                int componentEnd = Math.Min(volume.ComponentEnd, components.Length);
                if (volume.ResolutionStart >= resolutionEnd || volume.ComponentStart >= componentEnd || layerEnd <= 0) continue;
                switch (volume.Order) {
                    case 0: // LRCP
                        for (int layer = 0; layer < layerEnd; layer++)
                            for (int r = volume.ResolutionStart; r < resolutionEnd; r++)
                                for (int c = volume.ComponentStart; c < componentEnd; c++)
                                    if (!DecodeResolutionPackets(stream, components[c], r, layer, coding, budget)) return;
                        break;
                    case 1: // RLCP
                        for (int r = volume.ResolutionStart; r < resolutionEnd; r++)
                            for (int layer = 0; layer < layerEnd; layer++)
                                for (int c = volume.ComponentStart; c < componentEnd; c++)
                                    if (!DecodeResolutionPackets(stream, components[c], r, layer, coding, budget)) return;
                        break;
                    case 2: // RPCL
                        for (int r = volume.ResolutionStart; r < resolutionEnd; r++) {
                            var order = CollectPrecincts(components, volume.ComponentStart, componentEnd, r, r + 1, budget);
                            order.Sort(static (left, right) => ComparePositions(left, right, componentFirst: true));
                            if (!DecodeOrderedPrecincts(stream, components, order, layerEnd, coding, budget)) return;
                        }
                        break;
                    case 3: { // PCRL
                        var order = CollectPrecincts(components, volume.ComponentStart, componentEnd, volume.ResolutionStart, resolutionEnd, budget);
                        order.Sort(static (left, right) => ComparePositions(left, right, componentFirst: true));
                        if (!DecodeOrderedPrecincts(stream, components, order, layerEnd, coding, budget)) return;
                        break;
                    }
                    default: // CPRL
                        for (int c = volume.ComponentStart; c < componentEnd; c++) {
                            var order = CollectPrecincts(components, c, c + 1, volume.ResolutionStart, resolutionEnd, budget);
                            order.Sort(static (left, right) => ComparePositions(left, right, componentFirst: false));
                            if (!DecodeOrderedPrecincts(stream, components, order, layerEnd, coding, budget)) return;
                        }
                        break;
                }
            }
        } catch (PacketDataExhaustedException) {
            // Truncated packet data: keep every code-block contribution read so far (graceful degradation).
            stream.Exhausted = true;
        }
    }

    private readonly struct PrecinctReference {
        internal PrecinctReference(int component, int resolution, int index, long x, long y) {
            Component = component;
            Resolution = resolution;
            Index = index;
            X = x;
            Y = y;
        }

        internal int Component { get; }
        internal int Resolution { get; }
        internal int Index { get; }
        internal long X { get; }
        internal long Y { get; }
    }

    private static List<PrecinctReference> CollectPrecincts(TileComponent[] components, int componentStart, int componentEnd,
        int resolutionStart, int resolutionEnd, Budget budget) {
        var list = new List<PrecinctReference>();
        for (int c = componentStart; c < componentEnd; c++) {
            TileComponent component = components[c];
            int end = Math.Min(resolutionEnd, component.Resolutions.Length);
            for (int r = resolutionStart; r < end; r++) {
                Precinct[] precincts = component.Resolutions[r].Precincts;
                budget.Charge(precincts.Length);
                for (int k = 0; k < precincts.Length; k++) list.Add(new PrecinctReference(c, r, k, precincts[k].PositionX, precincts[k].PositionY));
            }
        }
        return list;
    }

    private static int ComparePositions(PrecinctReference left, PrecinctReference right, bool componentFirst) {
        int compare = left.Y.CompareTo(right.Y);
        if (compare != 0) return compare;
        compare = left.X.CompareTo(right.X);
        if (compare != 0) return compare;
        if (componentFirst) {
            compare = left.Component.CompareTo(right.Component);
            return compare != 0 ? compare : left.Resolution.CompareTo(right.Resolution);
        }
        compare = left.Resolution.CompareTo(right.Resolution);
        return compare != 0 ? compare : left.Component.CompareTo(right.Component);
    }

    private static bool DecodeOrderedPrecincts(TileStream stream, TileComponent[] components, List<PrecinctReference> order,
        int layerEnd, GlobalCoding coding, Budget budget) {
        foreach (PrecinctReference reference in order) {
            TileComponent component = components[reference.Component];
            Resolution resolution = component.Resolutions[reference.Resolution];
            Precinct precinct = resolution.Precincts[reference.Index];
            for (int layer = 0; layer < layerEnd; layer++) {
                budget.Charge(1);
                if (precinct.NextLayer != layer) continue;
                if (!DecodePacket(stream, resolution, precinct, layer, component.Coding.Style, budget)) return false;
            }
        }
        return true;
    }

    private static bool DecodeResolutionPackets(TileStream stream, TileComponent component, int r, int layer, GlobalCoding coding, Budget budget) {
        if (r >= component.Resolutions.Length) return true;
        Resolution resolution = component.Resolutions[r];
        foreach (Precinct precinct in resolution.Precincts) {
            budget.Charge(1);
            if (precinct.NextLayer != layer) continue;
            if (!DecodePacket(stream, resolution, precinct, layer, component.Coding.Style, budget)) return false;
        }
        return true;
    }

    // Packet header (B.10) and body.
    private static bool DecodePacket(TileStream stream, Resolution resolution, Precinct precinct, int layer, int style, Budget budget) {
        budget.Token.ThrowIfCancellationRequested();
        precinct.NextLayer = layer + 1;
        if (stream.Position >= stream.End && stream.PackedHeaders == null) return false;
        if (stream.Sop && stream.End - stream.Position >= 6 && stream.Data[stream.Position] == 0xFF && stream.Data[stream.Position + 1] == 0x91)
            stream.Position += 6;
        PacketHeaderReader reader = stream.PackedHeaders ?? stream.BodyHeaders!;
        if (stream.PackedHeaders == null) reader.Position = stream.Position;
        reader.StartPacket();
        List<(CodeSegment Segment, int Length)> pending = stream.Pending;
        pending.Clear();
        if (reader.ReadBit() != 0) {
            foreach (PrecinctBand band in precinct.Bands) {
                for (int index = 0; index < band.Blocks.Length; index++) {
                    CodeBlock block = band.Blocks[index];
                    bool included;
                    if (!block.Included) {
                        included = band.Inclusion!.Decode(reader, index, layer + 1);
                        if (included) {
                            block.ZeroPlanes = band.ZeroPlanes!.DecodeValue(reader, index, 74);
                            block.Included = true;
                        }
                    } else {
                        included = reader.ReadBit() != 0;
                    }
                    if (!included) continue;
                    int passes = ReadPassCount(reader);
                    while (reader.ReadBit() != 0) {
                        if (++block.LengthBits > 32) throw new Jpeg2000Exception("A JPEG 2000 code-block length indicator is invalid.");
                    }
                    if (block.Passes + passes > MaximumCodeBlockPasses) throw new Jpeg2000Exception("A JPEG 2000 code-block has too many coding passes.");
                    block.Segments ??= new List<CodeSegment>(1);
                    int remaining = passes;
                    while (remaining > 0) {
                        CodeSegment segment;
                        if (block.Segments.Count > 0 && block.Segments[block.Segments.Count - 1].Passes < block.Segments[block.Segments.Count - 1].MaximumPasses) {
                            segment = block.Segments[block.Segments.Count - 1];
                        } else {
                            budget.Reserve(64);
                            segment = new CodeSegment { MaximumPasses = MaximumSegmentPasses(style, block.Passes) };
                            block.Segments.Add(segment);
                        }
                        int take = Math.Min(remaining, segment.MaximumPasses - segment.Passes);
                        int bits = block.LengthBits + FloorLog2(take);
                        if (bits > 31) throw new Jpeg2000Exception("A JPEG 2000 codeword segment length is too large.");
                        pending.Add((segment, reader.ReadBits(bits)));
                        segment.Passes += take;
                        block.Passes += take;
                        remaining -= take;
                    }
                }
            }
        }
        reader.Align();
        if (stream.Eph) reader.SkipMarker(MarkerEph);
        if (stream.PackedHeaders == null) stream.Position = reader.Position;
        foreach ((CodeSegment segment, int length) in pending) {
            int available = Math.Min(length, stream.End - stream.Position);
            if (available > 0) {
                budget.Reserve(16);
                segment.Chunks.Add((stream.Position, available));
                segment.Length += available;
                stream.Position += available;
            }
            if (available < length) {
                stream.Exhausted = true;
                return false;
            }
        }
        return true;
    }

    // Number of coding passes (Table B.4).
    private static int ReadPassCount(PacketHeaderReader reader) {
        if (reader.ReadBit() == 0) return 1;
        if (reader.ReadBit() == 0) return 2;
        int value = reader.ReadBits(2);
        if (value != 3) return 3 + value;
        value = reader.ReadBits(5);
        if (value != 31) return 6 + value;
        return 37 + reader.ReadBits(7);
    }

    // Passes in a codeword segment starting at pass index `start` (D.4.1, Table D.9).
    private static int MaximumSegmentPasses(int style, int start) {
        if ((style & CodeBlockTerminateAll) != 0) return 1;
        if ((style & CodeBlockBypass) != 0) {
            if (start < 10) return 10 - start;
            return (start - 10) % 3 == 0 ? 2 : 1;
        }
        return int.MaxValue;
    }

    private static int FloorLog2(int value) {
        int result = 0;
        while ((value >>= 1) != 0) result++;
        return result;
    }
}
