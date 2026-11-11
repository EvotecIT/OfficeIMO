namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    private sealed class OutputChannel {
        internal ComponentPlane Plane = null!;
        internal int[]? PaletteColumn;
        internal int PalettePrecision;
        internal bool PaletteSigned;
        internal int Type; // cdef: 0 colour, 1 opacity, 2 premultiplied opacity
        internal int Association = 65535;

        internal int Precision => PaletteColumn != null ? PalettePrecision : Plane.Precision;
    }

    /// <summary>Output channels in interleaving order, their eight-bit lookups and the output samples.</summary>
    private sealed class OutputLayout {
        internal List<OutputChannel> Ordered = new();
        internal int ColourCount;
        internal OutputChannel? Alpha;
        internal bool UsedPalette;
        internal bool Container;
        internal byte[] Samples = Array.Empty<byte>();
        internal byte[][] Lookups = Array.Empty<byte[]>();
        internal int Width, Height;
    }

    // Plans decoding straight into the interleaved output: components on the reference grid are written through
    // their lookups as tiles are reconstructed, and only subsampled components keep a plane for upsampling. Returns
    // null when the channels cannot be resolved; whole planes are then decoded and the same error reported after.
    private static OutputLayout? PlanDirectOutput(Codestream codestream, ContainerHeader? header, ComponentPlane[] planes,
        OfficeJpeg2000DecodeOptions options, Budget budget) {
        long mark = budget.Reserved;
        OutputLayout layout;
        try {
            layout = ResolveOutput(codestream, header, planes, options, budget);
        } catch (Jpeg2000Exception) {
            budget.Release(budget.Reserved - mark);
            return null;
        }
        int count = layout.Ordered.Count;
        var direct = new List<int>[planes.Length];
        var stored = new bool[planes.Length];
        for (int c = 0; c < count; c++) {
            int index = Array.IndexOf(planes, layout.Ordered[c].Plane);
            if (codestream.SubsamplingX[index] == 1 && codestream.SubsamplingY[index] == 1) (direct[index] ??= new List<int>()).Add(c);
            else stored[index] = true;
        }
        long pixels = (long)codestream.Width * codestream.Height;
        for (int index = 0; index < planes.Length; index++) {
            ComponentPlane plane = planes[index];
            if (stored[index]) AllocatePlane(plane, budget);
            if (direct[index] == null) continue;
            plane.Output = layout.Samples;
            plane.OutputChannels = count;
            plane.OutputIndexes = direct[index].ToArray();
            plane.OutputLookups = Array.ConvertAll(plane.OutputIndexes, c => layout.Lookups[c]);
            // Missing tiles keep the level-shifted zero, the mid-grey of an unsigned component.
            foreach (int c in plane.OutputIndexes) {
                byte value = layout.Lookups[c][1 << (plane.Precision - 1)];
                byte[] samples = layout.Samples;
                for (long pixel = 0, target = c; pixel < pixels; pixel++, target += count) {
                    if ((pixel & 0xFFFF) == 0) budget.Token.ThrowIfCancellationRequested();
                    samples[target] = value;
                }
            }
        }
        return layout;
    }

    // Applies the JP2 palette, component mapping and channel definitions (I.5.3.4 to I.5.3.6) and allocates the
    // interleaved eight-bit output.
    private static OutputLayout ResolveOutput(Codestream codestream, ContainerHeader? header, ComponentPlane[] planes,
        OfficeJpeg2000DecodeOptions options, Budget budget) {
        bool container = options.ApplyContainerChannels && header != null;
        var channels = new List<OutputChannel>();
        bool usedPalette = false;
        if (container && header!.Palette != null && header.Mapping != null) {
            foreach ((int component, int type, int column) in header.Mapping) {
                if (component >= planes.Length) throw new Jpeg2000Exception("A JP2 component mapping refers to a missing component.");
                var channel = new OutputChannel { Plane = planes[component] };
                if (type == 1) {
                    if (column >= header.Palette.Columns.Length) throw new Jpeg2000Exception("A JP2 component mapping refers to a missing palette column.");
                    channel.PaletteColumn = header.Palette.Columns[column];
                    channel.PalettePrecision = header.Palette.Precisions[column];
                    channel.PaletteSigned = header.Palette.Signed[column];
                    usedPalette = true;
                } else if (type != 0) {
                    throw new Jpeg2000Exception("The JP2 component mapping type is invalid.");
                }
                channels.Add(channel);
            }
        } else {
            foreach (ComponentPlane plane in planes) channels.Add(new OutputChannel { Plane = plane });
        }
        if (channels.Count == 0) throw new Jpeg2000Exception("The JPEG 2000 image has no channels.");

        var colour = new List<OutputChannel>();
        OutputChannel? alpha = null;
        if (container && header!.Definitions != null && header.Definitions.Length > 0) {
            foreach ((int index, int type, int association) in header.Definitions) {
                if (index >= channels.Count) throw new Jpeg2000Exception("A JP2 channel definition refers to a missing channel.");
                channels[index].Type = type;
                channels[index].Association = association;
            }
            foreach (OutputChannel channel in channels) {
                if (channel.Type == 0 && channel.Association is > 0 and < 65535) colour.Add(channel);
                else if (channel.Type is 1 or 2 && alpha == null) alpha = channel;
            }
            colour.Sort(static (left, right) => left.Association.CompareTo(right.Association));
            if (colour.Count == 0) {
                foreach (OutputChannel channel in channels) if (channel.Type == 0) colour.Add(channel);
            }
        } else {
            colour.AddRange(channels);
        }
        if (colour.Count == 0) throw new Jpeg2000Exception("The JPEG 2000 image has no colour channels.");
        if (!options.ScaleToEightBits) {
            foreach (OutputChannel channel in colour)
                if (channel.Precision > 8) throw new Jpeg2000Exception("Unscaled JPEG 2000 samples wider than eight bits are not supported.");
        }

        int width = codestream.Width, height = codestream.Height;
        var ordered = new List<OutputChannel>(colour);
        if (alpha != null) ordered.Add(alpha);
        int count = ordered.Count;
        long total = (long)width * height * count;
        if (total > int.MaxValue) throw new Jpeg2000Exception("The JPEG 2000 image is too large to interleave.");
        budget.Reserve(total + (long)(width + height) * 4 * count);
        var samples = new byte[total];
        var lookups = new byte[count][];
        for (int c = 0; c < count; c++) lookups[c] = BuildLookup(ordered[c], options.ScaleToEightBits, budget);
        return new OutputLayout {
            Ordered = ordered, ColourCount = colour.Count, Alpha = alpha, UsedPalette = usedPalette, Container = container,
            Samples = samples, Lookups = lookups, Width = width, Height = height
        };
    }

    // Upsamples components kept as planes to the reference grid and interleaves them.
    private static void InterleavePlanes(OutputLayout layout, Codestream codestream, ComponentPlane[] planes, Budget budget) {
        int width = codestream.Width, height = codestream.Height;
        int count = layout.Ordered.Count;
        byte[] samples = layout.Samples;
        for (int c = 0; c < count; c++) {
            OutputChannel channel = layout.Ordered[c];
            ComponentPlane plane = channel.Plane;
            if (plane.Bytes == null && plane.Words == null) continue;
            int dx = codestream.SubsamplingX[Array.IndexOf(planes, plane)], dy = codestream.SubsamplingY[Array.IndexOf(planes, plane)];
            var columns = new int[width];
            for (int x = 0; x < width; x++) columns[x] = Clamp((codestream.XOsiz + x) / dx - plane.OriginX, plane.Width);
            byte[] lookup = layout.Lookups[c];
            for (int y = 0; y < height; y++) {
                if ((y & 63) == 0) budget.Token.ThrowIfCancellationRequested();
                int row = Clamp((codestream.YOsiz + y) / dy - plane.OriginY, plane.Height) * plane.Width;
                int target = y * width * count + c;
                if (plane.Bytes != null) {
                    byte[] source = plane.Bytes;
                    for (int x = 0; x < width; x++, target += count) samples[target] = lookup[source[row + columns[x]]];
                } else {
                    ushort[] source = plane.Words!;
                    for (int x = 0; x < width; x++, target += count) samples[target] = lookup[source[row + columns[x]]];
                }
            }
        }
    }

    private static OfficeJpeg2000Image CreateImage(OutputLayout layout, ContainerHeader? header, ComponentPlane[] planes,
        OfficeJpeg2000DecodeOptions options, Budget budget) {
        bool container = layout.Container;
        OutputChannel? alpha = layout.Alpha;
        byte[] samples = layout.Samples;
        var image = new OfficeJpeg2000Image(layout.Width, layout.Height, layout.ColourCount, alpha != null, alpha?.Type == 2, samples) {
            ColorSpace = container ? header!.ColorSpace : OfficeJpeg2000ColorSpace.Unspecified,
            EnumeratedColorSpace = container ? header!.EnumeratedColorSpace : -1,
            IccProfile = container ? header!.IccProfile : null,
            UsedPalette = layout.UsedPalette,
            Precisions = layout.Ordered.ConvertAll(static channel => channel.Precision).ToArray(),
            HasSignedComponents = Array.Exists(planes, static plane => plane.Signed)
        };
        if (image.ColorSpace == OfficeJpeg2000ColorSpace.Sycc && image.ColorChannels == 3 && options.ScaleToEightBits) {
            ConvertSyccToRgb(samples, image.Channels, budget.Token);
            image.ColorSpace = OfficeJpeg2000ColorSpace.Srgb;
        }
        return image;
    }

    private static int Clamp(int value, int count) => value < 0 ? 0 : value >= count ? count - 1 : value;

    /// <summary>Maps every stored sample value of a channel to its eight-bit output.</summary>
    private static byte[] BuildLookup(OutputChannel channel, bool scale, Budget budget) {
        ComponentPlane plane = channel.Plane;
        int entries = 1 << plane.Precision;
        budget.Reserve(entries);
        var lookup = new byte[entries];
        int precision = channel.Precision;
        long maximum = (1L << precision) - 1;
        for (int stored = 0; stored < entries; stored++) {
            int value;
            if (channel.PaletteColumn != null) {
                // Palette indices are unsigned component values (I.5.3.4).
                int index = plane.Signed ? stored - (1 << (plane.Precision - 1)) : stored;
                index = Clamp(index, channel.PaletteColumn.Length);
                value = channel.PaletteColumn[index];
                if (channel.PaletteSigned) value += 1 << (precision - 1);
            } else {
                value = stored;
            }
            if (!scale) {
                lookup[stored] = (byte)Math.Min(255, Math.Max(0, value));
                continue;
            }
            if (value < 0) value = 0;
            if (value > maximum) value = (int)maximum;
            lookup[stored] = precision == 8 ? (byte)value : (byte)((value * 255L + maximum / 2) / maximum);
        }
        return lookup;
    }

    // sYCC (IEC 61966-2-1 Amendment 1) to sRGB for eight-bit full-range samples.
    private static void ConvertSyccToRgb(byte[] samples, int channels, CancellationToken token) {
        for (int index = 0, pixel = 0; index + 2 < samples.Length; index += channels, pixel++) {
            if ((pixel & 0xFFFF) == 0) token.ThrowIfCancellationRequested();
            double luma = samples[index], blue = samples[index + 1] - 128.0, red = samples[index + 2] - 128.0;
            samples[index] = ToByte(luma + 1.402 * red);
            samples[index + 1] = ToByte(luma - 0.344136 * blue - 0.714136 * red);
            samples[index + 2] = ToByte(luma + 1.772 * blue);
        }
    }

    private static byte ToByte(double value) => value <= 0 ? (byte)0 : value >= 255 ? (byte)255 : (byte)(value + 0.5);
}
