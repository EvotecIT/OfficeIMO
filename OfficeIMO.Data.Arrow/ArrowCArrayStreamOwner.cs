using Apache.Arrow.C;
using Apache.Arrow.Ipc;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Data.Arrow;

/// <summary>
/// Owns an unmanaged Arrow C Data Interface <c>ArrowArrayStream</c> struct and
/// the managed stream exported through its callbacks.
/// </summary>
public sealed class ArrowCArrayStreamOwner : IDisposable {
    private const int Available = 0;
    private const int Leased = 1;
    private const int Consumed = 2;
    private const int Disposed = 3;
    private readonly ArrowCArrayStreamSafeHandle _handle;
    private int _state;

    private ArrowCArrayStreamOwner(ArrowCArrayStreamSafeHandle handle) {
        _handle = handle;
    }

    /// <summary>Gets whether the native stream was consumed or this owner was disposed.</summary>
    public bool IsDisposed => Volatile.Read(ref _state) >= Consumed;

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
    /// A stream owner grants one lease only; disposing that lease consumes the owner
    /// because native code may already have invoked the release callback.
    /// </remarks>
    public ArrowCArrayStreamLease AcquireLease() {
        int state = Interlocked.CompareExchange(ref _state, Leased, Available);
        if (state == Leased)
            throw new InvalidOperationException("The Arrow C stream already has an active native lease.");
        ObjectDisposedException.ThrowIf(state != Available, this);
        try {
            return new ArrowCArrayStreamLease(this, _handle);
        } catch {
            Interlocked.CompareExchange(ref _state, Available, Leased);
            throw;
        }
    }

    /// <summary>
    /// Imports this exported stream back into the managed Apache Arrow stream contract.
    /// </summary>
    /// <remarks>
    /// An import attempt consumes this owner. A successful import moves the native release
    /// callback to the returned stream. A failed import is also one-shot because the native
    /// importer may already have modified or released part of the callback table.
    /// </remarks>
    public IArrowArrayStream ImportArrayStream() {
        using ArrowCArrayStreamLease lease = AcquireLease();
        return lease.ImportArrayStream();
    }

    /// <summary>
    /// Stops accepting new leases and releases the unmanaged stream after active leases finish.
    /// </summary>
    public void Dispose() {
        if (Interlocked.Exchange(ref _state, Disposed) != Disposed) _handle.Dispose();
    }

    private bool BeginConsumption() {
        int state = Interlocked.CompareExchange(ref _state, Consumed, Leased);
        if (state == Leased) return true;
        if (state == Disposed) return false;
        throw new InvalidOperationException("The Arrow C stream is not owned by this lease.");
    }

    private void CompleteConsumption(bool stateChanged) {
        if (stateChanged) _handle.Dispose();
    }

    private void ConsumeLease() {
        if (Interlocked.CompareExchange(ref _state, Disposed, Leased) == Leased) _handle.Dispose();
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
        private const int Active = 0;
        private const int Importing = 1;
        private const int Consumed = 2;
        private const int LeaseDisposed = 3;
        private readonly ArrowCArrayStreamOwner _owner;
        private readonly object _sync = new();
        private ArrowCArrayStreamSafeHandle? _handle;
        private readonly nint _address;
        private int _state;

        internal ArrowCArrayStreamLease(ArrowCArrayStreamOwner owner, ArrowCArrayStreamSafeHandle handle) {
            _owner = owner;
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
                ObjectDisposedException.ThrowIf(_handle == null || Volatile.Read(ref _state) == LeaseDisposed, this);
                if (Volatile.Read(ref _state) != Active)
                    throw new InvalidOperationException("The Arrow C stream callback ownership has already been consumed.");
                return _address;
            }
        }

        /// <summary>Returns the native pointer while this lease is active.</summary>
        public unsafe CArrowArrayStream* DangerousGetPointer() =>
            (CArrowArrayStream*)Address;

        /// <summary>
        /// Imports the leased native stream into Apache Arrow's managed stream contract.
        /// </summary>
        public unsafe IArrowArrayStream ImportArrayStream() {
            lock (_sync) {
                ObjectDisposedException.ThrowIf(_handle == null, this);
                if (Interlocked.CompareExchange(ref _state, Importing, Active) != Active)
                    throw new InvalidOperationException("The Arrow C stream can be imported only once.");

                bool ownerStateChanged = false;
                try {
                    ownerStateChanged = _owner.BeginConsumption();
                    return CArrowArrayStreamImporter.ImportArrayStream((CArrowArrayStream*)_address);
                } finally {
                    Volatile.Write(ref _state, Consumed);
                    _owner.CompleteConsumption(ownerStateChanged);
                }
            }
        }

        /// <summary>Releases this lease.</summary>
        public void Dispose() {
            lock (_sync) {
                int priorState = Interlocked.Exchange(ref _state, LeaseDisposed);
                ArrowCArrayStreamSafeHandle? handle = Interlocked.Exchange(ref _handle, null);
                if (handle != null) handle.DangerousRelease();
                if (priorState == Active) _owner.ConsumeLease();
            }
            GC.SuppressFinalize(this);
        }

        /// <summary>Releases the native reference if the lease was abandoned.</summary>
        ~ArrowCArrayStreamLease() {
            int priorState = Interlocked.Exchange(ref _state, LeaseDisposed);
            ArrowCArrayStreamSafeHandle? handle = Interlocked.Exchange(ref _handle, null);
            if (handle != null) handle.DangerousRelease();
            if (priorState == Active) _owner.ConsumeLease();
        }
    }
}
