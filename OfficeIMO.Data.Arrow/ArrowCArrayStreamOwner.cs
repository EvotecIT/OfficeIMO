using Apache.Arrow.C;
using Apache.Arrow.Ipc;

namespace OfficeIMO.Data.Arrow;

/// <summary>
/// Owns an unmanaged Arrow C Data Interface <c>ArrowArrayStream</c> struct and
/// the managed stream exported through its callbacks.
/// </summary>
public sealed class ArrowCArrayStreamOwner : IDisposable {
    private nint _address;

    private ArrowCArrayStreamOwner(nint address) {
        _address = address;
    }

    /// <summary>Gets the unmanaged stream address, or zero after disposal.</summary>
    public nint Address => Volatile.Read(ref _address);

    /// <summary>Gets whether this owner has released its unmanaged stream.</summary>
    public bool IsDisposed => Address == 0;

    internal static unsafe ArrowCArrayStreamOwner Export(IArrowArrayStream stream) {
        ArgumentNullException.ThrowIfNull(stream);
        CArrowArrayStream* pointer = CArrowArrayStream.Create();
        try {
            CArrowArrayStreamExporter.ExportArrayStream(stream, pointer);
            return new ArrowCArrayStreamOwner((nint)pointer);
        } catch {
            CArrowArrayStream.Free(pointer);
            throw;
        }
    }

    /// <summary>
    /// Returns the native pointer for an immediate interop call.
    /// </summary>
    /// <remarks>
    /// The pointer is valid only while this owner remains undisposed. Native code may
    /// invoke the stream's release callback, but must not free the struct allocation.
    /// </remarks>
    public unsafe CArrowArrayStream* DangerousGetPointer() {
        nint address = Address;
        ObjectDisposedException.ThrowIf(address == 0, this);
        return (CArrowArrayStream*)address;
    }

    /// <summary>Releases the exported stream and frees its unmanaged struct.</summary>
    public void Dispose() {
        Release();
        GC.SuppressFinalize(this);
    }

    ~ArrowCArrayStreamOwner() {
        Release();
    }

    private unsafe void Release() {
        nint address = Interlocked.Exchange(ref _address, 0);
        if (address != 0) CArrowArrayStream.Free((CArrowArrayStream*)address);
    }
}
