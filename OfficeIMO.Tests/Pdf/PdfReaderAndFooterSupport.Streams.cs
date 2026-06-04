using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReaderAndFooterRegressionTests {

    private static byte[] BuildPdfWithPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        return BuildSingleStreamPdf(compressedBytes, $"/Filter /FlateDecode /DecodeParms << /Predictor 12 /Columns {streamBytes.Length} >>");
    }

    private static byte[] BuildPdfWithAscii85AndPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(compressedBytes));
        return BuildSingleStreamPdf(encodedBytes, $"/Filter [/ASCII85Decode /FlateDecode] /DecodeParms [null << /Predictor 12 /Columns {streamBytes.Length} >>]");
    }

    private static byte[] BuildPdfWithIndirectPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        return BuildSingleStreamPdfWithExtraObjects(
            compressedBytes,
            "/Filter /FlateDecode /DecodeParms 6 0 R",
            "6 0 obj",
            $"<< /Predictor 12 /Columns {streamBytes.Length} >>",
            "endobj");
    }

    private static byte[] BuildPdfWithAscii85AndIndirectPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(compressedBytes));
        return BuildSingleStreamPdfWithExtraObjects(
            encodedBytes,
            "/Filter [/ASCII85Decode /FlateDecode] /DecodeParms [null 6 0 R]",
            "6 0 obj",
            $"<< /Predictor 12 /Columns {streamBytes.Length} >>",
            "endobj");
    }

    private static byte[] BuildPdfWithTiffPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeTiffPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        return BuildSingleStreamPdf(compressedBytes, $"/Filter /FlateDecode /DecodeParms << /Predictor 2 /Columns {streamBytes.Length} >>");
    }

    private static byte[] BuildPdfWithIndirectFilterNameEncodedStream(string streamContent) {
        byte[] compressedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdfWithExtraObjects(
            compressedBytes,
            "/Filter 6 0 R",
            "6 0 obj",
            "/FlateDecode",
            "endobj");
    }

    private static byte[] BuildPdfWithIndirectFilterAndDecodeParmsArrayObjects(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(compressedBytes));
        return BuildSingleStreamPdfWithExtraObjects(
            encodedBytes,
            "/Filter 6 0 R /DecodeParms 7 0 R",
            "6 0 obj",
            "[/ASCII85Decode /FlateDecode]",
            "endobj",
            "7 0 obj",
            "[null 8 0 R]",
            "endobj",
            "8 0 obj",
            $"<< /Predictor 12 /Columns {streamBytes.Length} >>",
            "endobj");
    }

    private static byte[] BuildPdfWithIndirectLengthStreamContainingEndstreamLiteral() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello endstream marker) Tj\nET\n";
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        return BuildSingleStreamPdfWithExtraObjects(
            streamBytes,
            "/Length 6 0 R",
            "6 0 obj",
            streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "endobj");
    }


    private static byte[] BuildPdfWithFlateCompressedStream(string streamContent) {
        byte[] compressedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdf(compressedBytes, "/Filter /FlateDecode");
    }

    private static byte[] BuildPdfWithAsciiHexEncodedStream(string streamContent) {
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAsciiHex(Encoding.ASCII.GetBytes(streamContent)));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /ASCIIHexDecode");
    }

    private static byte[] BuildPdfWithAscii85EncodedStream(string streamContent) {
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(Encoding.ASCII.GetBytes(streamContent)));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /ASCII85Decode");
    }

    private static byte[] BuildPdfWithAscii85AndFlateEncodedStream(string streamContent) {
        byte[] flatedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(flatedBytes));
        return BuildSingleStreamPdf(encodedBytes, "/Filter [/ASCII85Decode /FlateDecode]");
    }

    private static byte[] BuildPdfWithAsciiHexAndFlateEncodedStream(string streamContent) {
        byte[] flatedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAsciiHex(flatedBytes));
        return BuildSingleStreamPdf(encodedBytes, "/Filter [/AHx /Fl]");
    }

    private static byte[] BuildPdfWithRunLengthEncodedStream(string streamContent) {
        byte[] encodedBytes = EncodeRunLength(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /RunLengthDecode");
    }

    private static byte[] BuildPdfWithAscii85AndRunLengthEncodedStream(string streamContent) {
        byte[] runLengthBytes = EncodeRunLength(Encoding.ASCII.GetBytes(streamContent));
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(runLengthBytes));
        return BuildSingleStreamPdf(encodedBytes, "/Filter [/A85 /RL]");
    }

    private static byte[] BuildPdfWithLzwEncodedStream(string streamContent) {
        byte[] encodedBytes = EncodeLzw(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /LZWDecode");
    }

    private static byte[] BuildPdfWithAscii85AndLzwEncodedStream(string streamContent, int earlyChange) {
        byte[] lzwBytes = EncodeLzw(Encoding.ASCII.GetBytes(streamContent), earlyChange);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(lzwBytes));
        return BuildSingleStreamPdf(encodedBytes, $"/Filter [/A85 /LZW] /DecodeParms [null << /EarlyChange {earlyChange} >>]");
    }

    private static byte[] BuildPdfWithLzwPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] encodedBytes = EncodeLzw(predictedBytes);
        return BuildSingleStreamPdf(encodedBytes, $"/Filter /LZWDecode /DecodeParms << /Predictor 12 /Columns {streamBytes.Length} >>");
    }


    private static byte[] CompressWithDeflate(byte[] input) {
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(input, 0, input.Length);
        }

        return output.ToArray();
    }

    private static byte[] EncodeUpPredictedRows(byte[] input) {
        var output = new byte[input.Length + 1];
        output[0] = 2;
        Buffer.BlockCopy(input, 0, output, 1, input.Length);
        return output;
    }

    private static byte[] EncodeTiffPredictedRows(byte[] input) {
        if (input.Length == 0) {
            return Array.Empty<byte>();
        }

        var output = new byte[input.Length];
        output[0] = input[0];
        for (int i = 1; i < input.Length; i++) {
            output[i] = unchecked((byte)(input[i] - input[i - 1]));
        }

        return output;
    }

    private static byte[] EncodeRunLength(byte[] input) {
        using var output = new MemoryStream();
        int index = 0;
        while (index < input.Length) {
            int chunkLength = Math.Min(128, input.Length - index);
            output.WriteByte((byte)(chunkLength - 1));
            output.Write(input, index, chunkLength);
            index += chunkLength;
        }

        output.WriteByte(128);
        return output.ToArray();
    }

    private static byte[] EncodeLzw(byte[] input, int earlyChange = 1) {
        earlyChange = earlyChange == 0 ? 0 : 1;
        var dictionary = new Dictionary<string, int>();
        for (int i = 0; i < 256; i++) {
            dictionary[Convert.ToBase64String(new[] { (byte)i })] = i;
        }

        using var output = new MemoryStream();
        var writer = new LzwBitWriter(output);
        int nextCode = 258;
        int codeSize = 9;
        writer.WriteBits(256, codeSize);
        var current = new List<byte>();

        foreach (byte value in input) {
            var candidate = new List<byte>(current) { value };
            string candidateKey = Convert.ToBase64String(candidate.ToArray());
            if (dictionary.ContainsKey(candidateKey)) {
                current = candidate;
                continue;
            }

            writer.WriteBits(dictionary[Convert.ToBase64String(current.ToArray())], codeSize);
            if (nextCode <= 4095) {
                dictionary[candidateKey] = nextCode++;
                if (codeSize < 12 && nextCode + earlyChange >= (1 << codeSize)) {
                    codeSize++;
                }
            }

            current = new List<byte> { value };
        }

        if (current.Count > 0) {
            writer.WriteBits(dictionary[Convert.ToBase64String(current.ToArray())], codeSize);
        }

        writer.WriteBits(257, codeSize);
        writer.Flush();
        return output.ToArray();
    }

    private static string EncodeAsciiHex(byte[] input) {
        var sb = new StringBuilder(input.Length * 2 + 1);
        for (int i = 0; i < input.Length; i++) {
            sb.Append(input[i].ToString("X2"));
        }
        sb.Append('>');
        return sb.ToString();
    }

    private static string EncodeAscii85(byte[] input) {
        var sb = new StringBuilder((input.Length * 5 / 4) + 4);
        int index = 0;
        while (index + 4 <= input.Length) {
            uint value =
                ((uint)input[index] << 24) |
                ((uint)input[index + 1] << 16) |
                ((uint)input[index + 2] << 8) |
                input[index + 3];

            if (value == 0) {
                sb.Append('z');
            } else {
                AppendAscii85Tuple(sb, value, 5);
            }

            index += 4;
        }

        int remaining = input.Length - index;
        if (remaining > 0) {
            uint value = 0;
            for (int i = 0; i < remaining; i++) {
                value |= (uint)input[index + i] << (24 - (8 * i));
            }

            AppendAscii85Tuple(sb, value, remaining + 1);
        }

        sb.Append("~>");
        return sb.ToString();
    }

    private static void AppendAscii85Tuple(StringBuilder sb, uint value, int count) {
        char[] encoded = new char[5];
        for (int i = 4; i >= 0; i--) {
            encoded[i] = (char)((value % 85) + '!');
            value /= 85;
        }

        for (int i = 0; i < count; i++) {
            sb.Append(encoded[i]);
        }
    }

    private sealed class LzwBitWriter {
        private readonly Stream _stream;
        private int _currentByte;
        private int _bitCount;

        public LzwBitWriter(Stream stream) {
            _stream = stream;
        }

        public void WriteBits(int value, int bitCount) {
            for (int i = bitCount - 1; i >= 0; i--) {
                _currentByte = (_currentByte << 1) | ((value >> i) & 1);
                _bitCount++;
                if (_bitCount == 8) {
                    _stream.WriteByte((byte)_currentByte);
                    _currentByte = 0;
                    _bitCount = 0;
                }
            }
        }

        public void Flush() {
            if (_bitCount == 0) {
                return;
            }

            _stream.WriteByte((byte)(_currentByte << (8 - _bitCount)));
            _currentByte = 0;
            _bitCount = 0;
        }
    }

}
