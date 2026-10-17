using System;

namespace OfficeIMO.Archives;

/// <summary>Thread-safe cumulative budget shared explicitly by related archive operations.
/// Pass CreateNested() to a child container. Reservations are never refunded; this bounds cumulative work,
/// not live retained memory. Reusing this object deliberately continues the same operation.</summary>
public sealed class ArchiveOperationBudget {
    private readonly State _state;
    private readonly int _depth;
    /// <summary>Creates a root budget and snapshots its nonnegative limits.</summary>
    public ArchiveOperationBudget(ArchiveOperationLimits? limits = null) { _state = new State(limits ?? new ArchiveOperationLimits()); }
    private ArchiveOperationBudget(State state, int depth) { _state = state; _depth = depth; }
    /// <summary>Returns a child view sharing all counters. Child input admission enforces the depth limit.</summary>
    public ArchiveOperationBudget CreateNested() => new(_state, checked(_depth + 1));
    /// <summary>Returns a detached snapshot of the entire tree's charges.</summary>
    public ArchiveResourceUsage Usage { get { lock (_state) return new(_state.Input, _state.Entries, _state.Expanded,
        _state.Selected, _state.Work, _state.Depth); } }
    internal void AdmitInput(long bytes) { lock (_state) {
        Check(_depth, _state.DepthLimit, "MaxDepth");
        Add(ref _state.Input, bytes, _state.InputLimit, "MaxInputBytes");
        _state.Depth = Math.Max(_state.Depth, _depth);
    } }
    internal void AdmitEntries(long count) { lock (_state) Add(ref _state.Entries, count, _state.EntryLimit, "MaxEntryCount"); }
    internal void Select(long bytes) { lock (_state) Add(ref _state.Selected, bytes, _state.SelectedLimit, "MaxSelectedBytes"); }
    internal void ReadZip(long selected, long expanded, long work) { lock (_state) {
        CheckZipRead(selected, expanded, work);
        _state.Selected += selected; _state.Expanded += expanded; _state.Work += work;
    } }
    internal void CheckZipRead(long selected, long expanded, long work) { lock (_state) {
        Check(selected, _state.SelectedLimit - _state.Selected, "MaxSelectedBytes", _state.SelectedLimit);
        Check(expanded, _state.ExpandedLimit - _state.Expanded, "MaxExpandedBytes", _state.ExpandedLimit);
        Check(work, _state.WorkLimit - _state.Work, "MaxDecompressionWorkBytes", _state.WorkLimit);
    } }
    internal void Decode(long expanded, long work) { lock (_state) {
        Check(expanded, _state.ExpandedLimit - _state.Expanded, "MaxExpandedBytes", _state.ExpandedLimit);
        Check(work, _state.WorkLimit - _state.Work, "MaxDecompressionWorkBytes", _state.WorkLimit);
        _state.Expanded += expanded; _state.Work += work;
    } }
    private static void Add(ref long used, long amount, long maximum, string name) {
        Check(amount, maximum - used, name, maximum); used += amount;
    }
    private static void Check(long amount, long remaining, string name, long? maximum = null) {
        if (amount < 0) throw new ArgumentOutOfRangeException(nameof(amount));
        if (amount > remaining) throw ArchiveLimits.Exceeded("Operation." + name, maximum ?? remaining);
    }
    private sealed class State {
        internal State(ArchiveOperationLimits limits) {
            InputLimit = Valid(limits.MaxInputBytes); EntryLimit = Valid(limits.MaxEntryCount);
            ExpandedLimit = Valid(limits.MaxExpandedBytes); SelectedLimit = Valid(limits.MaxSelectedBytes);
            WorkLimit = Valid(limits.MaxDecompressionWorkBytes); DepthLimit = checked((int)Valid(limits.MaxDepth));
        }
        private static long Valid(long value) => value >= 0 ? value : throw new ArgumentOutOfRangeException(nameof(ArchiveOperationLimits));
        internal readonly long InputLimit, EntryLimit, ExpandedLimit, SelectedLimit, WorkLimit;
        internal readonly int DepthLimit;
        internal long Input, Entries, Expanded, Selected, Work;
        internal int Depth;
    }
}
