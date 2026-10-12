#nullable enable

#if NET10_0_OR_GREATER
using System;
using System.Data;
using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Data;

public static partial class DataReaderMappingExtensions {
    /// <summary>Projects unread rows with a caller factory into models that may contain borrowed spans.</summary>
    /// <typeparam name="T">The mapped model, including a ref struct.</typeparam>
    /// <param name="reader">Reader positioned before the next row to project.</param>
    /// <param name="factory">Creates one model from the current row. Called once for each successful advance.</param>
    /// <param name="cancellationToken">Cancels between rows and before invoking the factory.</param>
    /// <returns>A synchronous, forward-only pattern enumerable.</returns>
    /// <remarks>
    /// Uses the caller-factory contract of RowsAs: the factory controls construction, column access,
    /// and conversion through the provider's getters. It does not perform automatic property mapping.
    /// Resolve ordinals before enumeration when they should be reused across rows. Getter and factory
    /// failures propagate unchanged. The caller owns the reader, including after failure or early exit.
    /// A model containing borrowed field spans is valid only until the reader advances to another row
    /// or result, or closes or is disposed. Consume it within the loop and do not advance the reader
    /// externally or enumerate the same reader concurrently. Copy fields that must survive advancement.
    /// Current is cached, so accessing it repeatedly does not call the factory again. Cancellation
    /// cannot interrupt a running synchronous read or factory. This API provides no asynchronous,
    /// parallel, or materialized result contract.
    /// </remarks>
    public static BorrowedRowEnumerable<T> RowsAsBorrowed<T>(
        this DbDataReader reader,
        Func<IDataRecord, T> factory,
        CancellationToken cancellationToken = default) where T : allows ref struct {
        ArgumentNullException.ThrowIfNull(reader);
        ArgumentNullException.ThrowIfNull(factory);
        return new BorrowedRowEnumerable<T>(reader, factory, cancellationToken);
    }
}

/// <summary>A synchronous factory-mapped sequence whose models may borrow the current reader's fields.</summary>
/// <typeparam name="T">The mapped model, including a ref struct.</typeparam>
/// <remarks>
/// Enumerates remaining unread rows without taking ownership of the reader. Models and field spans
/// must be consumed before the next reader advance, result change, close, or disposal.
/// This pattern enumerable does not expose an interface or a materialization API.
/// </remarks>
public readonly ref struct BorrowedRowEnumerable<T> where T : allows ref struct {
    private readonly DbDataReader _reader;
    private readonly Func<IDataRecord, T> _factory;
    private readonly CancellationToken _cancellationToken;

    internal BorrowedRowEnumerable(DbDataReader reader, Func<IDataRecord, T> factory, CancellationToken cancellationToken) {
        _reader = reader;
        _factory = factory;
        _cancellationToken = cancellationToken;
    }

    /// <summary>Creates an enumerator over the reader's remaining unread rows.</summary>
    public Enumerator GetEnumerator() => new(_reader, _factory, _cancellationToken);

    /// <summary>Retains one mapped model until the next advance or disposal.</summary>
    /// <remarks>Disposal ends this enumeration and leaves the caller-owned reader open.</remarks>
    public ref struct Enumerator {
        private DbDataReader? _reader;
        private Func<IDataRecord, T>? _factory;
        private CancellationToken _cancellationToken;
        private T _current;
        private bool _started;
        private bool _hasCurrent;
        private bool _completed;

        internal Enumerator(DbDataReader reader, Func<IDataRecord, T> factory, CancellationToken cancellationToken) {
            _reader = reader;
            _factory = factory;
            _cancellationToken = cancellationToken;
            _current = default!;
            _started = false;
            _hasCurrent = false;
            _completed = false;
        }

        /// <summary>Gets the model produced by the most recent successful advance.</summary>
        /// <exception cref="InvalidOperationException">No successful advance is current, or enumeration has ended.</exception>
        public T Current => _hasCurrent
            ? _current
            : throw new InvalidOperationException("The enumerator is not positioned on a mapped row.");

        /// <summary>Advances the source and invokes the factory exactly once for the next row.</summary>
        /// <returns>True when a mapped row is current; otherwise false.</returns>
        /// <remarks>An exception ends this enumeration without disposing the reader.</remarks>
        public bool MoveNext() {
            if (_completed) return false;
            _hasCurrent = false;
            _current = default!;
            try {
                _cancellationToken.ThrowIfCancellationRequested();
                if (_reader is null || _factory is null) {
                    throw new InvalidOperationException("The enumerable has not been initialized.");
                }
                if (!_started) {
                    _started = true;
                    if (_reader.FieldCount == 0) {
                        Dispose();
                        return false;
                    }
                }
                if (!_reader.Read()) {
                    Dispose();
                    return false;
                }
                _cancellationToken.ThrowIfCancellationRequested();
                _current = _factory(_reader);
                _hasCurrent = true;
                return true;
            } catch {
                Dispose();
                throw;
            }
        }

        /// <summary>Clears the current model and ends enumeration without closing the reader.</summary>
        public void Dispose() {
            _hasCurrent = false;
            _current = default!;
            _completed = true;
            _factory = null;
            _reader = null;
            _cancellationToken = default;
        }
    }
}
#endif
