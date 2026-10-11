namespace OfficeIMO.DjVu;

internal sealed partial class Iw44Decoder {
    private readonly DjVuReadBudget _budget;
    private Iw44Plane[]? _planes;
    private int _serial, _delay, _slices;
    private bool _halfChroma;
    internal int Width { get; private set; }
    internal int Height { get; private set; }
    internal long RetainedBytes => _planes == null ? 0 : _planes.Sum(p => p.Coefficients.LongLength * 4);
    internal Iw44Decoder(DjVuReadBudget budget) => _budget = budget;

    internal void Read(DjVuChunk chunk) {
        _budget.Cancellation.ThrowIfCancellationRequested();
        byte[] data = chunk.Source;
        int position = chunk.Offset, end = position + chunk.Length;
        if (end - position < 2 || data[position++] != _serial++) throw new InvalidDataException("Invalid IW44 progressive chunk sequence.");
        int slices = data[position++];
        if (slices > _budget.Options.MaxIw44Slices - _slices)
            throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxIw44Slices));
        _slices += slices;
        if (_planes == null) {
            if (end - position < 7) throw new InvalidDataException("Truncated IW44 header.");
            int major = data[position++], minor = data[position++];
            if ((major & 127) != 1 || minor != 2) throw new NotSupportedException("The managed decoder supports IW44 codec version 1.2.");
            Width = DjVuBinary.U16(data, position); Height = DjVuBinary.U16(data, position + 2); position += 4;
            int chroma = data[position++];
            _delay = chroma & 127;
            _halfChroma = (chroma & 128) == 0;
            if (Width <= 0 || Height <= 0) throw new InvalidDataException("Empty IW44 image.");
            if ((long)Width * Height > _budget.Options.MaxPagePixels) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxPagePixels));
            int blocks = checked(((Width + 31) / 32) * ((Height + 31) / 32));
            int components = (major & 128) != 0 ? 1 : 3;
            _budget.WorkingBytes(blocks * 1024L * components * 4 + (long)Width * Height * (components * 4 + 4));
            _planes = new Iw44Plane[components];
            for (int i = 0; i < components; i++) _planes[i] = new Iw44Plane(blocks, _budget);
        }
        long work = _planes[0].CoefficientSamples(slices);
        if (_planes.Length == 3) {
            int chromaSlices = Math.Max(0, slices - _delay);
            work += _planes[1].CoefficientSamples(chromaSlices) + _planes[2].CoefficientSamples(chromaSlices);
        }
        _budget.Iw44CoefficientSamples(work);
        // An all-MPS early slice can have no entropy bytes. ZP defines implicit trailing one bits.
        var arithmetic = new ZpDecoder(data, position, end - position, _budget.Cancellation);
        for (int i = 0; i < slices; i++) {
            _planes[0].Slice(arithmetic);
            if (_delay == 0 && _planes.Length == 3) { _planes[1].Slice(arithmetic); _planes[2].Slice(arithmetic); }
            if (_delay > 0) _delay--;
        }
    }

    internal byte[] Reconstruct() {
        if (_planes == null) throw new InvalidDataException("Missing IW44 image data.");
        _budget.WorkingBytes(RetainedBytes + (long)Width * Height * (_planes.Length * 4 + 3));
        var channels = new int[_planes.Length][];
        for (int i = 0; i < channels.Length; i++) channels[i] = ReconstructPlane(_planes[i], i != 0 && _halfChroma ? 2 : 1);
        var rgb = new byte[checked(Width * Height * 3)];
        for (int y = 0; y < Height; y++) {
            _budget.Cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < Width; x++) {
                int input = y * Width + x, output = ((Height - 1 - y) * Width + x) * 3;
                int luminance = Math.Max(-128, Math.Min(127, (channels[0][input] + 32) >> 6));
                if (channels.Length == 1) { byte gray = (byte)(127 - luminance); rgb[output] = rgb[output + 1] = rgb[output + 2] = gray; }
                else {
                    // Half-resolution chroma uses one reconstructed sample per 2x2
                    // native grid cell, including the final odd-sized row/column.
                    int chromaInput = _halfChroma ? (y & ~1) * Width + (x & ~1) : input;
                    int cb = Math.Max(-128, Math.Min(127, (channels[1][chromaInput] + 32) >> 6));
                    int cr = Math.Max(-128, Math.Min(127, (channels[2][chromaInput] + 32) >> 6));
                    int yy = luminance + 128;
                    rgb[output] = Clamp(yy + cr + (cr >> 1));
                    rgb[output + 1] = Clamp(yy - (cb >> 2) - ((3 * cr) >> 2));
                    rgb[output + 2] = Clamp(yy + cb * 2 - (cb >> 2));
                }
            }
        }
        return rgb;
    }
    private static byte Clamp(int value) => (byte)Math.Max(0, Math.Min(255, value));
}
