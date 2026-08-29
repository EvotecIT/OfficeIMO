using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>Decodes already-validated XML bytes using their declaration while allowing a BOM to take precedence.</summary>
internal static class OfficeXmlTextEncoding {
    internal static string Decode(byte[] bytes, string? declaredEncoding) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));

        Encoding encoding = new UTF8Encoding(false, true);
        if (!string.IsNullOrWhiteSpace(declaredEncoding)) {
            try {
                encoding = Encoding.GetEncoding(
                    declaredEncoding,
                    EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback);
            } catch (ArgumentException exception) {
                throw new InvalidDataException($"Unsupported XML encoding '{declaredEncoding}'.", exception);
            } catch (NotSupportedException exception) {
                throw new InvalidDataException($"Unsupported XML encoding '{declaredEncoding}'.", exception);
            }
        }

        using var memory = new MemoryStream(bytes, writable: false);
        using var reader = new StreamReader(
            memory,
            encoding,
            detectEncodingFromByteOrderMarks: true,
            bufferSize: 1024,
            leaveOpen: false);
        return reader.ReadToEnd();
    }
}
