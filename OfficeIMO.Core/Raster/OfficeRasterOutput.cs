using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal static class OfficeRasterOutput {
    internal static void EnsureWritable(Stream destination) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) {
            throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        }
    }
}
