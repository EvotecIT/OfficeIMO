namespace OfficeIMO.Pdf.Filters;

internal static class LzwDecoder {
    private const int ClearCode = 256;
    private const int EndOfDataCode = 257;
    private const int FirstAvailableCode = 258;
    private const int MaximumCodeValue = 4095;
    private const int MaximumCodeSize = 12;

    public static byte[] Decode(byte[] data, int earlyChange = 1) {
        if (!TryDecode(data, int.MaxValue, out byte[] output, earlyChange)) {
            throw new InvalidOperationException("LZWDecode output exceeded the configured limit.");
        }

        return output;
    }

    public static bool TryDecode(byte[] data, int maxOutputBytes, out byte[] outputBytes, int earlyChange = 1) {
        if (data == null || data.Length == 0) {
            outputBytes = Array.Empty<byte>();
            return true;
        }

        earlyChange = earlyChange == 0 ? 0 : 1;
        var reader = new BitReader(data);
        using var output = new MemoryStream();
        var dictionary = CreateInitialDictionary();
        int nextCode = FirstAvailableCode;
        int codeSize = 9;
        byte[]? previous = null;

        while (true) {
            int code = reader.ReadBits(codeSize);
            if (code < 0) {
                break;
            }

            if (code == ClearCode) {
                dictionary = CreateInitialDictionary();
                nextCode = FirstAvailableCode;
                codeSize = 9;
                previous = null;
                continue;
            }

            if (code == EndOfDataCode) {
                break;
            }

            byte[] entry;
            if (code < dictionary.Count && dictionary[code] is not null) {
                entry = dictionary[code]!;
            } else if (code == nextCode && previous is not null) {
                entry = AppendByte(previous, previous[0]);
            } else {
                throw new FormatException("Invalid LZWDecode code.");
            }

            if (output.Length + entry.Length > maxOutputBytes) {
                outputBytes = Array.Empty<byte>();
                return false;
            }

            output.Write(entry, 0, entry.Length);

            if (previous is not null && nextCode <= MaximumCodeValue) {
                AddEntry(dictionary, nextCode, AppendByte(previous, entry[0]));
                nextCode++;
                if (codeSize < MaximumCodeSize && nextCode + earlyChange >= (1 << codeSize)) {
                    codeSize++;
                }
            }

            previous = entry;
        }

        outputBytes = output.ToArray();
        return true;
    }

    private static List<byte[]?> CreateInitialDictionary() {
        var dictionary = new List<byte[]?>(MaximumCodeValue + 1);
        for (int i = 0; i < 256; i++) {
            dictionary.Add(new[] { (byte)i });
        }

        dictionary.Add(null);
        dictionary.Add(null);
        return dictionary;
    }

    private static void AddEntry(List<byte[]?> dictionary, int code, byte[] entry) {
        if (code == dictionary.Count) {
            dictionary.Add(entry);
        } else if (code < dictionary.Count) {
            dictionary[code] = entry;
        } else {
            throw new FormatException("Invalid LZWDecode dictionary state.");
        }
    }

    private static byte[] AppendByte(byte[] bytes, byte value) {
        var result = new byte[bytes.Length + 1];
        Buffer.BlockCopy(bytes, 0, result, 0, bytes.Length);
        result[result.Length - 1] = value;
        return result;
    }

    private sealed class BitReader {
        private readonly byte[] _data;
        private int _bitOffset;

        public BitReader(byte[] data) {
            _data = data;
        }

        public int ReadBits(int bitCount) {
            if (bitCount <= 0 || _bitOffset + bitCount > _data.Length * 8) {
                return -1;
            }

            int value = 0;
            for (int i = 0; i < bitCount; i++) {
                int absoluteBit = _bitOffset + i;
                int currentByte = _data[absoluteBit / 8];
                int bit = (currentByte >> (7 - (absoluteBit % 8))) & 1;
                value = (value << 1) | bit;
            }

            _bitOffset += bitCount;
            return value;
        }
    }
}
