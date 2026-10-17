using System;
using System.IO;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives {

    /// <summary>Decodes declared ZIP names deterministically without registering a global code-page provider.</summary>
    internal static class ArchiveZipNames {
        internal static string Decode(OfficeArchiveSafety.ZipCentralDirectoryRecord record, int? legacyCodePage) {
            byte[] bytes = record.NameBytes;
            try {
                if ((record.Flags & 0x0800) != 0) return new UTF8Encoding(false, true).GetString(bytes);
                bool ascii = true;
                foreach (byte value in bytes) if (value > 127) { ascii = false; break; }
                if (ascii) return Encoding.ASCII.GetString(bytes);
                if (!legacyCodePage.HasValue) throw new NotSupportedException("Unflagged non-ASCII ZIP names require an explicit ZipLegacyCodePage.");
                int codePage = legacyCodePage.Value;
                if (codePage == 65001) return new UTF8Encoding(false, true).GetString(bytes);
    #if NET8_0_OR_GREATER
                Encoding? encoding = CodePagesEncodingProvider.Instance.GetEncoding(codePage,
                    EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                encoding ??= Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                return encoding.GetString(bytes);
    #else
                if (codePage == 1252 || codePage == 1250 || codePage == 437 || codePage == 850 || codePage == 10000)
                    return OfficeLegacySingleByteEncoding.Decode(bytes, 0, bytes.Length, codePage);
                return Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback).GetString(bytes);
    #endif
            } catch (DecoderFallbackException error) { throw new InvalidDataException("Invalid ZIP filename bytes for the selected encoding.", error); }
            catch (ArgumentException error) { throw new NotSupportedException("Unsupported ZIP filename encoding.", error); }
        }
    }
}
