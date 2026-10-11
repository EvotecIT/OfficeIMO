using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

internal sealed partial class MsiPackageReader {
    private void LoadStrings() {
        byte[] pool = ReadMetadata("_StringPool");
        byte[] data = ReadMetadata("_StringData");
        uint header = UInt32(pool, 0);
        _stringWidth = (header & 0x80000000U) == 0 ? 2 : 3;
        int codePage = checked((int)(header & 0x7fffffff));
        var strings = new List<string?> { null };
        int dataOffset = 0;
        for (int offset = 4; offset < pool.Length;) {
            _token.ThrowIfCancellationRequested();
            if (strings.Count > _limits.Rows) throw new InvalidDataException("MSI string count exceeds the configured metadata row limit.");
            int length = UInt16(pool, offset);
            int references = UInt16(pool, offset + 2);
            offset += 4;
            if (length == 0 && references != 0) {
                uint extended = UInt32(pool, offset);
                if (extended > int.MaxValue) throw new InvalidDataException("MSI string length exceeds supported bounds.");
                length = (int)extended; offset += 4;
            }
            if (length > data.Length - dataOffset) throw new InvalidDataException("MSI string pool exceeds its data stream.");
            strings.Add(DecodeString(data, dataOffset, length, codePage));
            dataOffset += length;
        }
        if (dataOffset != data.Length) throw new InvalidDataException("MSI string data contains unaccounted bytes.");
        _strings = strings.ToArray();
    }
    private string DecodeString(byte[] data, int offset, int length, int codePage) {
        try {
            if (codePage == 0) return Encoding.GetEncoding(20127, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback).GetString(data, offset, length);
            if (codePage == 65001) return new UTF8Encoding(false, true).GetString(data, offset, length);
#if NET8_0_OR_GREATER
            Encoding? encoding = CodePagesEncodingProvider.Instance.GetEncoding(codePage,
                EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            if (encoding != null) return encoding.GetString(data, offset, length);
            throw new NotSupportedException("The MSI string code page is unsupported: " + codePage);
#else
            if (codePage == 1252 || codePage == 1250 || codePage == 437 || codePage == 850 || codePage == 10000)
                return OfficeLegacySingleByteEncoding.Decode(data, offset, length, codePage);
            return Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback).GetString(data, offset, length);
#endif
        } catch (DecoderFallbackException exception) { throw new InvalidDataException("MSI string data is invalid for its declared code page.", exception); }
        catch (ArgumentException exception) { throw new NotSupportedException("The MSI string code page is unsupported: " + codePage, exception); }
    }
}
