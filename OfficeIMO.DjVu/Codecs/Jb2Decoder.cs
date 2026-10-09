namespace OfficeIMO.DjVu;

// DjVu v3 appendix 2. No decoder implementation from a comparison engine is used.
internal sealed class Jb2Decoder {
    private const int Minimum = -262143, Maximum = 262142;
    private readonly DjVuReadBudget _budget;
    private readonly ZpDecoder _arithmetic;
    private readonly Jb2NumberDecoder _numbers;
    private readonly byte[] _direct = new byte[1024], _refinement = new byte[2048];
    private readonly List<Jb2Bitmap> _library = new List<Jb2Bitmap>();
    private readonly List<Jb2Placement> _placements = new List<Jb2Placement>();
    private readonly IReadOnlyList<Jb2Bitmap>? _inherited;
    private readonly long _inheritedBytes;
    private byte _eventualRefinement, _offsetType;
    private int _width, _height, _symbols, _firstX, _firstY, _previousRight, _lineCount;
    private readonly int[] _baselines = new int[3];
    private long _bitmapBytes;

    internal Jb2Decoder(DjVuChunk chunk, DjVuReadBudget budget, IReadOnlyList<Jb2Bitmap>? inherited = null) {
        if (chunk.Length == 0) throw new InvalidDataException("Empty JB2 stream.");
        _budget = budget; _inherited = inherited;
        _inheritedBytes = inherited == null ? 0 : inherited.Distinct().Sum(b => b.Pixels.LongLength + 32L) + inherited.Count * 8L;
        _arithmetic = new ZpDecoder(chunk.Source, chunk.Offset, chunk.Length, budget.Cancellation);
        _numbers = new Jb2NumberDecoder(_arithmetic, budget, () => _inheritedBytes + BitmapBuffers);
    }

    internal Jb2Image Decode(bool dictionary = false) {
        bool started = false, inheritedSeen = false;
        int records = 0;
        while (true) {
            _budget.Cancellation.ThrowIfCancellationRequested();
            if (++records > (long)_budget.Options.MaxSymbols + _budget.Options.MaxSymbolPlacements + _budget.Options.MaxChunks)
                throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxSymbols));
            int record = Number(0, 0, 11);
            if (record == 9) {
                if (started) _numbers.Reset();
                else {
                    if (inheritedSeen) throw new InvalidDataException("Repeated JB2 shared dictionary requirement.");
                    inheritedSeen = true;
                    int count = Number(15, 0, Maximum);
                    if (count > _budget.Options.MaxSymbols) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxSymbols));
                    if (_inherited == null || _inherited.Count != count) throw new InvalidDataException("JB2 shared dictionary size does not match.");
                    _library.AddRange(_inherited);
                    _symbols = count;
                    CheckBuffers(0);
                }
                continue;
            }
            if (record == 0) {
                if (started) throw new InvalidDataException("Repeated JB2 start record.");
                started = true;
                _width = Number(1, 0, Maximum); _height = Number(1, 0, Maximum);
                if (dictionary ? _width != 0 || _height != 0 : _width <= 0 || _height <= 0) throw new InvalidDataException("Invalid JB2 image dimensions.");
                if ((long)_width * _height > _budget.Options.MaxPagePixels) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxPagePixels));
                if (_arithmetic.Bit(ref _eventualRefinement) != 0) throw new NotSupportedException("JB2 eventual image refinement is unsupported.");
                _firstX = _previousRight = -1;
                _firstY = _height - 1;
                continue;
            }
            if (!started) throw new InvalidDataException("JB2 data precedes its start record.");
            if (record == 11) return new Jb2Image(_width, _height, _library, _placements);
            if (record == 10) {
                int length = Number(13, 0, Maximum);
                for (int i = 0; i < length; i++) { if ((i & 4095) == 0) _budget.Cancellation.ThrowIfCancellationRequested(); Number(14, 0, 255); }
                continue;
            }
            if (dictionary && record != 2 && record != 5) throw new InvalidDataException("Image record in JB2 dictionary.");
            Jb2Bitmap bitmap;
            if (record >= 4 && record <= 7) {
                if (_library.Count == 0) throw new InvalidDataException("JB2 match has no library.");
                bitmap = _library[Number(2, 0, _library.Count - 1)];
                if (record != 7) {
                    int width = checked(bitmap.Width + Number(5, Minimum, Maximum));
                    int height = checked(bitmap.Height + Number(6, Minimum, Maximum));
                    bitmap = DecodeBitmap(width, height, bitmap);
                }
            } else if (record >= 1 && record <= 3 || record == 8) {
                int width = Number(3, 0, Maximum), height = Number(4, 0, Maximum);
                bitmap = DecodeBitmap(width, height, null);
            } else throw new InvalidDataException("Unknown JB2 record.");
            if (record != 2 && record != 5) {
                if (_placements.Count >= _budget.Options.MaxSymbolPlacements) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxSymbolPlacements));
                int x, y;
                if (record == 8) {
                    x = Number(7, 1, _width) - 1;
                    y = Number(8, 1, _height) - bitmap.Height;
                } else Relative(bitmap.Width, bitmap.Height, out x, out y);
                _placements.Add(new Jb2Placement(bitmap, x, y));
                CheckBuffers(0);
            }
            if (record == 1 || record == 2 || record == 4 || record == 5) {
                CheckBuffers(bitmap.Pixels.LongLength);
                var trimmed = bitmap.Trim(_budget.Cancellation);
                if (!ReferenceEquals(bitmap, trimmed)) {
                    _bitmapBytes += trimmed.Pixels.Length;
                    if (record == 2 || record == 5) _bitmapBytes -= bitmap.Pixels.Length;
                }
                _library.Add(trimmed);
            }
        }
    }

    private int Number(int context, int low, int high) => _numbers.Read(context, low, high);

    private long BitmapBuffers => _bitmapBytes + _placements.Count * 40L + _library.Count * 32L + 4096;
    private void CheckBuffers(long additional) => _budget.WorkingBytes(_inheritedBytes + BitmapBuffers + additional + _numbers.RetainedBytes);

    private Jb2Bitmap DecodeBitmap(int width, int height, Jb2Bitmap? reference) {
        if (width < 0 || height < 0 || width > Maximum || height > Maximum) throw new InvalidDataException("Invalid JB2 symbol dimensions.");
        if (++_symbols > _budget.Options.MaxSymbols) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxSymbols));
        long count = (long)width * height;
        if (count > int.MaxValue) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxCodecBytes));
        _budget.Jb2DecodedSamples(count);
        CheckBuffers(count);
        var bitmap = new Jb2Bitmap(width, height, new byte[(int)count]);
        _bitmapBytes += count;
        int dx = reference == null ? 0 : (reference.Width - 1) / 2 - (width - 1) / 2;
        int dy = reference == null ? 0 : (reference.Height - 1) / 2 - (height - 1) / 2;
        for (int y = height - 1; y >= 0; y--) {
            _budget.Cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                int context;
                if (reference == null) {
                    context = bitmap.At(x - 1, y + 2) << 9 | bitmap.At(x, y + 2) << 8 | bitmap.At(x + 1, y + 2) << 7
                        | bitmap.At(x - 2, y + 1) << 6 | bitmap.At(x - 1, y + 1) << 5 | bitmap.At(x, y + 1) << 4
                        | bitmap.At(x + 1, y + 1) << 3 | bitmap.At(x + 2, y + 1) << 2 | bitmap.At(x - 2, y) << 1 | bitmap.At(x - 1, y);
                    bitmap.Pixels[y * width + x] = (byte)_arithmetic.Bit(ref _direct[context]);
                } else {
                    int rx = x + dx, ry = y + dy;
                    context = bitmap.At(x - 1, y + 1) << 10 | bitmap.At(x, y + 1) << 9 | bitmap.At(x + 1, y + 1) << 8 | bitmap.At(x - 1, y) << 7
                        | reference.At(rx, ry + 1) << 6 | reference.At(rx - 1, ry) << 5 | reference.At(rx, ry) << 4 | reference.At(rx + 1, ry) << 3
                        | reference.At(rx - 1, ry - 1) << 2 | reference.At(rx, ry - 1) << 1 | reference.At(rx + 1, ry - 1);
                    bitmap.Pixels[y * width + x] = (byte)_arithmetic.Bit(ref _refinement[context]);
                }
            }
        }
        return bitmap;
    }

    private void Relative(int width, int height, out int x, out int y) {
        checked {
            if (_arithmetic.Bit(ref _offsetType) != 0) {
                x = _firstX + Number(11, Minimum, Maximum);
                int top = _firstY + Number(12, Minimum, Maximum);
                y = top - height + 1;
                _firstX = x; _firstY = y;
                _baselines[0] = _baselines[1] = _baselines[2] = y;
                _lineCount = 1;
            } else {
                x = _previousRight + Number(9, Minimum, Maximum);
                int median = _baselines[0] + _baselines[1] + _baselines[2] - Math.Min(_baselines[0], Math.Min(_baselines[1], _baselines[2])) - Math.Max(_baselines[0], Math.Max(_baselines[1], _baselines[2]));
                y = median + Number(10, Minimum, Maximum);
                _baselines[_lineCount++ % 3] = y;
            }
            _previousRight = x + width - 1;
        }
    }
}
