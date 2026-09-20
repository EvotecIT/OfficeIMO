using Apache.Arrow.C;
using Apache.Arrow.Ipc;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Data.Arrow;

/// <summary>
/// Owns an unmanaged Arrow C Data Interface <c>ArrowArrayStream</c> struct and
/// the managed stream exported through its callbacks.
/// </summary>
public sealed class ArrowCArrayStreamOwner : IDisposable {
    private readonly ArrowCArrayStreamSafeHandle _handle;
    private int _disposed;

    private ArrowCArrayStreamOwner(ArrowCArrayStreamSafeHandle handle) {
        _handle = handle;
    }

    /// <summary>Gets whether this owner no longer accepts new native leases.</summary>
    public bool IsDisposed => Volatile.Read(ref _disposed) != 0;

    internal static unsafe ArrowCArrayStreamOwner Export(IArrowArrayStream stream) {
        ArgumentNullException.ThrowIfNull(stream);
        CArrowArrayStream* pointer = CArrowArrayStream.Create();
        try {
            CArrowArrayStreamExporter.ExportArrayStream(stream, pointer);
            return new ArrowCArrayStreamOwner(new ArrowCArrayStreamSafeHandle((nint)pointer));
        } catch {
            CArrowArrayStream.Free(pointer);
            throw;
        }
    }

    /// <summary>
    /// Acquires a fail-safe lease for one native consumer call or callback sequence.
    /// </summary>
    /// <remarks>
    /// The native address is exposed only by the returned lease. Disposing this owner
    /// cannot free the stream struct until every lease has been disposed or finalized.
    /// Native code may invoke the stream release callback, but it must not free the
    /// struct allocation. Keep the lease alive for the complete native call sequence.
    /// </remarks>
    public ArrowCArrayStreamLease AcquireLease() {
        ObjectDisposedException.ThrowIf(IsDisposed, this);
        return new ArrowCArrayStreamLease(_handle);
    }

    /// <summary>
    /// Imports this exported stream back into the managed Apache Arrow stream contract.
    /// </summary>
    /// <remarks>
    /// A successful import moves the native release callback to the returned stream and
    /// disposes this owner. If import fails, this owner remains available for retry or
    /// disposal.
    /// </remarks>
    public IArrowArrayStream ImportArrayStream() {
        using ArrowCArrayStreamLease lease = AcquireLease();
        IArrowArrayStream imported = lease.ImportArrayStream();
        Dispose();
        return imported;
    }

    /// <summary>
    /// Stops accepting new leases and releases the unmanaged stream after active leases finish.
    /// </summary>
    public void Dispose() {
        if (Interlocked.Exchange(ref _disposed, 1) == 0) _handle.Dispose();
    }

    internal sealed class ArrowCArrayStreamSafeHandle : SafeHandleZeroOrMinusOneIsInvalid {
        internal ArrowCArrayStreamSafeHandle(nint address) : base(ownsHandle: true) {
            SetHandle(address);
        }

        protected override unsafe bool ReleaseHandle() {
            CArrowArrayStream.Free((CArrowArrayStream*)handle);
            return true;
        }
    }

    /// <summary>
    /// Pins one exported Arrow C stream allocation while a native consumer uses it.
    /// </summary>
    public sealed class ArrowCArrayStreamLease : IDisposable {
        private ArrowCArrayStreamSafeHandle? _handle;
        private readonly nint _address;

        internal ArrowCArrayStreamLease(ArrowCArrayStreamSafeHandle handle) {
            bool addedRef = false;
            try {
                handle.DangerousAddRef(ref addedRef);
                ObjectDisposedException.ThrowIf(handle.IsInvalid, handle);
                _address = handle.DangerousGetHandle();
                _handle = handle;
            } catch {
                if (addedRef) handle.DangerousRelease();
                throw;
            }
        }

        /// <summary>Gets the native stream address while this lease is active.</summary>
        public nint Address {
            get {
                ObjectDisposedException.ThrowIf(_handle == null, this);
                return _address;
            }
        }

        /// <summary>Returns the native pointer while this lease is active.</summary>
        public unsafe CArrowArrayStream* DangerousGetPointer() =>
            (CArrowArrayStream*)Address;

        /// <summary>
        /// Imports the leased native stream into Apache Arrow's managed stream contract.
        /// </summary>
        public unsafe IArrowArrayStream ImportArrayStream() =>
            CArrowArrayStreamImporter.ImportArrayStream(DangerousGetPointer());

        /// <summary>Releases this lease.</summary>
        public void Dispose() {
            ArrowCArrayStreamSafeHandle? handle = Interlocked.Exchange(ref _handle, null);
            if (handle != null) handle.DangerousRelease();
            GC.SuppressFinalize(this);
        }

        /// <summary>Releases the native reference if the lease was abandoned.</summary>
        ~ArrowCArrayStreamLease() {
            ArrowCArrayStreamSafeHandle? handle = Interlocked.Exchange(ref _handle, null);
            if (handle != null) handle.DangerousRelease();
        }
    }
}
