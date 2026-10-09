using System;

namespace OfficeIMO.Drawing;

/// <summary>Accounts for live encoding buffers and the transient peak during their growth.</summary>
internal sealed class OfficeVp8EncodingMemory {
    private long _reservedBytes;

    internal OfficeVp8EncodingMemory(long reservedBytes) {
        Reserve(reservedBytes);
    }

    internal void Reserve(long bytes) {
        if (bytes < 0 || bytes > OfficeRasterGuards.MaximumDecodedBytes - _reservedBytes) {
            throw new ArgumentException("Lossy WebP encoding exceeds the managed working-set limit.");
        }
        _reservedBytes += bytes;
    }

    internal void ReplaceBuffer(int oldLength, int newLength) {
        // Array.Resize temporarily retains both backing arrays. Check that peak before allocation.
        Reserve(newLength + 24L);
        _reservedBytes -= oldLength + 24L;
    }

    internal void EnsureAdditionalBytes(long bytes) {
        if (bytes < 0 || bytes > OfficeRasterGuards.MaximumDecodedBytes - _reservedBytes) {
            throw new ArgumentException("Lossy WebP encoding exceeds the managed working-set limit.");
        }
    }
}
