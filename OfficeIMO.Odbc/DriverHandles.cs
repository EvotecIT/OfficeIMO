using System.Collections.Concurrent;
using System.Data.Common;

namespace OfficeIMO.Odbc;

internal abstract class DriverHandle(short kind, nint parent) {
    internal readonly object Sync = new();
    internal short Kind { get; } = kind;
    internal nint Parent { get; } = parent;
    internal List<DriverDiagnostic> Diagnostics { get; } = [];
    internal bool Freed;
    internal short Fail(string state, string message) { Diagnostics.Add(new(state, message)); return -1; }
}

internal sealed record DriverDiagnostic(string State, string Message);
internal sealed class EnvironmentHandle() : DriverHandle(1, 0);
internal sealed class ConnectionHandle(nint parent) : DriverHandle(2, parent) {
    internal WorkbookSnapshot? Snapshot;
}
internal sealed class DescriptorHandle(nint parent, StatementHandle statement, int attribute) : DriverHandle(4, parent) {
    internal StatementHandle Statement { get; } = statement;
    internal int Attribute { get; } = attribute;
}
internal sealed class StatementHandle(nint parent) : DriverHandle(3, parent) {
    internal Dictionary<int, nint> Descriptors { get; } = [];
    internal string? Sql;
    internal DbDataReader? Reader;
    internal CancellationTokenSource? Cancellation;
    internal Dictionary<ushort, ColumnBinding> Bindings { get; } = [];
    internal Dictionary<ushort, ParameterBinding> Parameters { get; } = [];
    internal Dictionary<ushort, int> Offsets { get; } = [];
    internal bool OnRow;
    internal long Fetched;
    internal int ParameterCount;
    internal long MaximumRows;
    internal int TimeoutSeconds = 30;
    internal nint RowsFetchedPointer;
    internal nint RowStatusPointer;
    internal void CloseCursor() {
        try { Reader?.Dispose(); }
        finally {
            Reader = null;
            Cancellation?.Dispose();
            Cancellation = null;
            Offsets.Clear();
            OnRow = false;
            Fetched = 0;
        }
    }
}
internal sealed record ColumnBinding(short Type, nint Buffer, nint Length, nint Indicator);
internal sealed record ParameterBinding(short Type, nint Buffer, nint Length, nint Indicator);

/// <summary>Opaque monotonic IDs prevent native callers from dereferencing managed or stale pointers.</summary>
internal static class DriverHandles {
    internal static readonly ConcurrentDictionary<nint, DriverHandle> All = new();
    private static long _next;
    internal static short Allocate(short kind, nint parent, out nint id) {
        id = 0;
        lock (All) {
            if (All.Count + (kind == 3 ? 5 : 1) > 4096) return -1;
            DriverHandle? owner = null;
            if (kind != 1 && (!All.TryGetValue(parent, out owner) || owner.Freed || owner.Kind != kind - 1)) return -2;
            if (kind == 2 && All.Values.Count(handle => handle.Kind == 2) >= 32) return owner!.Fail("HY001", "Connection limit reached.");
            DriverHandle? handle = kind switch {
                1 => new EnvironmentHandle(), 2 => new ConnectionHandle(parent), 3 => new StatementHandle(parent), _ => null
            };
            if (handle is null) return owner?.Fail("HYC00", "Explicit descriptors are not supported.") ?? (short)-1;
            id = (nint)Interlocked.Increment(ref _next);
            All[id] = handle;
            if (handle is StatementHandle statement) {
                for (int attribute = 10010; attribute <= 10013; attribute++) {
                    nint descriptor = (nint)Interlocked.Increment(ref _next);
                    All[descriptor] = new DescriptorHandle(id, statement, attribute);
                    statement.Descriptors[attribute] = descriptor;
                }
            }
            return 0;
        }
    }

    internal static short With<T>(nint id, Func<T, short> operation) where T : DriverHandle {
        if (!All.TryGetValue(id, out DriverHandle? handle) || handle is not T typed) return -2;
        // Serialize driver calls; cancellation deliberately bypasses this gate.
        lock (All) lock (handle.Sync) {
            if (handle.Freed) return -2;
            handle.Diagnostics.Clear();
            try { return operation(typed); }
            catch (OperationCanceledException) { return handle.Fail("HY008", "Query canceled or its execution deadline expired."); }
            catch (NotSupportedException exception) { return handle.Fail("HYC00", exception.Message); }
            catch (ArgumentException exception) { return handle.Fail("HY090", exception.Message); }
            catch (Exception exception) { return handle.Fail("HY000", "OfficeIMO could not complete the operation: " + exception.GetBaseException().Message); }
        }
    }
}
