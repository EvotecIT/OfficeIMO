using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeArchiveSafety {
    // Framework searches all earlier bytes for a locator when the classic count is 65,535.
    // Reject that ambiguous shape before ZipArchive can materialize a different directory.
    private static bool HasEarlierZip64Locator(Stream source, long archiveOffset, long endOffset, CancellationToken token) {
        var buffer = new byte[81923];
        long position = 0;
        int carry = 0;
        source.Position = archiveOffset;
        while (position < endOffset) {
            token.ThrowIfCancellationRequested();
            int read = source.Read(buffer, carry, (int)Math.Min(81920, endOffset - position));
            if (read == 0) throw new InvalidDataException("ZIP data ended during locator validation.");
            int count = carry + read;
            if (HasEarlierZip64Locator(buffer, 0, count)) return true;
            carry = Math.Min(3, count);
            Array.Copy(buffer, count - carry, buffer, 0, carry);
            position += read;
        }
        return false;
    }

    private static bool HasEarlierZip64Locator(byte[] bytes, int start, int end) {
        for (int index = start; index <= end - 4; index++)
            if (ReadZipUInt32(bytes, index) == Zip64EndOfCentralDirectoryLocatorSignature) return true;
        return false;
    }
}
