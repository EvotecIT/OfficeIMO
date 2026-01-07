using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace OfficeIMO.Shared;

/// <summary>
/// MemoryStream that suppresses disposal so OpenXml packages can close without losing the buffer.
/// </summary>
internal sealed class NonDisposingMemoryStream : MemoryStream
{
    public NonDisposingMemoryStream(int capacity) : base(capacity)
    {
    }

    public NonDisposingMemoryStream(byte[] buffer) : base(buffer)
    {
    }

    [SuppressMessage("Usage", "CA2215:Dispose methods should call base class dispose",
        Justification = "Suppress disposal so the buffer remains accessible after OpenXml closes the stream.")]
    protected override void Dispose(bool disposing)
    {
        // Suppress disposal so the buffer remains accessible after OpenXml closes the stream.
    }

    public void DisposeUnderlying()
    {
        base.Dispose(true);
    }
}
