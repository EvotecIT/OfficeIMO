# TODO: Excel Processing Refactor – Parallel vs Sequential Control

## Problem Summary

The Excel fluent API fails to open files when certain operations are combined (AutoFilter, ConditionalFormatting, AutoFit, Column operations). Investigation revealed:

* All operations (parallel and sequential) use the same `ReaderWriterLockSlim` instance.
* Locks are acquired even for non-concurrent operations (unnecessary overhead).
* The parallel threshold (1000 cells) is rarely reached in typical fluent API usage.
* Locking issues affect both parallel and sequential code paths equally.
* Nested lock acquisitions with inconsistent recursion handling cause deadlocks.

## Root Cause

1. **Unnecessary locking**: Sequential operations don’t need thread synchronization.
2. **Shared lock instance everywhere**: Both parallel and sequential paths took the same `_lock`.
3. **Inconsistent lock handling**: Methods sometimes checked `IsWriteLockHeld`, sometimes not.
4. **Multiple structural modifications**: AutoFilter + AutoFit + ConditionalFormatting conflicted.

---

## Architecture Design

### 1) Execution Policy (Document/Sheet/Call)

```csharp
public enum ExecutionMode
{
    Automatic,   // Default: decide by thresholds
    Sequential,  // Force single-threaded, no locks
    Parallel     // Compute in parallel; single serialized apply
}

public sealed class ExecutionPolicy
{
    public ExecutionMode Mode { get; set; } = ExecutionMode.Automatic;

    /// <summary>Default threshold above which Automatic switches to Parallel.</summary>
    public int ParallelThreshold { get; set; } = 10_000;

    /// <summary>Per-operation thresholds (names: "CellValues", "InsertObjects", "AutoFitColumns", ...)</summary>
    public Dictionary<string, int> OperationThresholds { get; } = new(StringComparer.Ordinal);

    /// <summary>Optional cap for parallel compute phase.</summary>
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>Structured diagnostics (operation, items, decided mode).</summary>
    public Action<string,int,ExecutionMode>? OnDecision { get; set; }

    internal ExecutionMode Decide(string op, int count)
    {
        var thr = OperationThresholds.TryGetValue(op, out var v) ? v : ParallelThreshold;
        var decided = count > thr ? ExecutionMode.Parallel : ExecutionMode.Sequential;
        OnDecision?.Invoke(op, count, decided);
        return decided;
    }
}
```

**Recommended defaults (can be overridden):**

* `CellValues`: 10,000
* `InsertObjects`: 1,000
* `AutoFitColumns`, `AutoFitRows`: 2,000
* `ConditionalFormatting` (any variant): 2,000

### 2) Document & Sheet Integration

```csharp
public class ExcelDocument : IDisposable
{
    public ExecutionPolicy Execution { get; } = new();

    // Allocated only when an operation actually needs a serialized apply stage.
    internal ReaderWriterLockSlim? _lock;

    internal ReaderWriterLockSlim EnsureLock()
        => _lock ??= new ReaderWriterLockSlim(); // default: NoRecursion

    public void Dispose() => _lock?.Dispose();
}

public partial class ExcelSheet
{
    private readonly ExcelDocument _document;

    /// <summary>Null = inherit from document.</summary>
    public ExecutionPolicy? ExecutionOverride { get; set; }

    internal ExecutionPolicy EffectiveExecution => ExecutionOverride ?? _document.Execution;

    // Keep this PUBLIC if you want power users to control lock skipping globally,
    // or make it INTERNAL and let the fluent builder use it behind the scenes.
    public NoLockContext BeginNoLock() => new();

    public sealed class NoLockContext : IDisposable
    {
        private readonly IDisposable _scope;
        internal NoLockContext() => _scope = Locking.EnterNoLockScope();
        public void Dispose() => _scope.Dispose();
    }
}
```

### 3) Locking Infrastructure (centralized)

```csharp
internal static class Locking
{
    private static readonly AsyncLocal<bool> _noLockScope = new();

    public static IDisposable EnterNoLockScope()
    {
        var prev = _noLockScope.Value;
        _noLockScope.Value = true;
        return new Scope(() => _noLockScope.Value = prev);
    }

    public static bool IsNoLock => _noLockScope.Value;

    private sealed class Scope : IDisposable
    {
        private readonly Action _onDispose;
        public Scope(Action onDispose) => _onDispose = onDispose;
        public void Dispose() => _onDispose();
    }

    /// <summary>Serialize the short apply-to-DOM stage only.</summary>
    public static void ExecuteWrite(ReaderWriterLockSlim? lck, Action apply)
    {
        if (IsNoLock || lck is null) { apply(); return; }
        lck.EnterWriteLock();
        try { apply(); }
        finally { lck.ExitWriteLock(); }
    }
}
```

### 4) Core Execution Helper (compute outside lock, apply inside)

> **Important fix:** compute runs **without** locks; only the apply stage is serialized.

```csharp
private void ExecuteWithPolicy(
    string opName,
    int itemCount,
    ExecutionMode? overrideMode,
    Action sequentialCore,                // single-threaded path (no locks)
    Action? computeParallel = null,       // parallelizable compute (no DOM)
    Action? applySequential = null,       // serialized DOM apply
    CancellationToken ct = default)
{
    var policy = EffectiveExecution;
    var mode = overrideMode ?? policy.Mode;
    if (mode == ExecutionMode.Automatic)
        mode = policy.Decide(opName, itemCount);

    if (mode == ExecutionMode.Sequential || computeParallel is null || applySequential is null)
    {
        using (Locking.EnterNoLockScope())
            sequentialCore();
        return;
    }

    // Parallel: compute without lock
    var po = new ParallelOptions { CancellationToken = ct };
    if (policy.MaxDegreeOfParallelism is int dop && dop > 0)
        po.MaxDegreeOfParallelism = dop;

    computeParallel();

    // Apply once, serialized
    Locking.ExecuteWrite(_document.EnsureLock(), applySequential);
}
```

---

## Method Patterns

### A) Cells (no public \*NoLock methods—minimal surface)

```csharp
// Single cell: trivially sequential
public void CellValue(int row, int column, object value, ExecutionMode? mode = null)
{
    // No need to involve the helper for single items
    CellValueCore(row, column, value);
}

public void CellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null, CancellationToken ct = default)
{
    var list = cells as IList<(int Row, int Column, object Value)> ?? cells.ToList();
    if (list.Count == 0) return;

    // PREPARED BUFFERS
    var prepared = new (int Row, int Col, string Val, string Type)[list.Count];

    ExecuteWithPolicy(
        opName: "CellValues",
        itemCount: list.Count,
        overrideMode: mode,
        sequentialCore: () =>
        {
            for (int i = 0; i < list.Count; i++)
            {
                var (r, c, v) = list[i];
                CellValueCore(r, c, v);
            }
        },
        computeParallel: () =>
        {
            Parallel.For(0, list.Count, new ParallelOptions {
                CancellationToken = ct,
                MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
            }, i =>
            {
                var (r, c, obj) = list[i];
                var (val, type) = CoerceForCell(obj); // must NOT touch SharedStrings/Styles
                prepared[i] = (r, c, val, type);
            });

            // SharedStrings/Styles planning: collect distincts here (thread-safe),
            // DO NOT mutate OpenXML parts in compute. See planner below.
        },
        applySequential: () =>
        {
            // Resolve shared strings/styles indexes now, then write cells
            _sharedStringPlanner.ApplyAndFixup(prepared, this); // see planner pattern
            for (int i = 0; i < prepared.Length; i++)
            {
                var p = prepared[i];
                CellValueCorePrepared(p.Row, p.Col, p.Val, p.Type);
            }
        },
        ct: ct
    );
}

// Core implementation: single source of truth (no locks here)
private void CellValueCore(int row, int column, object value)
{
    var (val, type) = CoerceForCell(value); // MUST NOT mutate OpenXML parts
    // If type == "SharedString", index resolution is deferred to apply phase via planner
    SimulatedDomSet(row, column, val, type);
}

private void CellValueCorePrepared(int row, int column, string val, string type)
{
    SimulatedDomSet(row, column, val, type);
}

// Compute-only coercion (no OpenXML mutations)
private (string Val, string Type) CoerceForCell(object value)
{
    // Return raw text + logical type hint ("Number","Date","SharedString",...)
    // SharedString text stays as text here; actual index is assigned in apply phase by planner.
    switch (value)
    {
        case null: return (string.Empty, "String");
        case string s: return (s, "SharedString");
        case int i: return (i.ToString(), "Number");
        case double d: return (d.ToString(CultureInfo.InvariantCulture), "Number");
        case DateTime dt: return (dt.ToOADate().ToString(CultureInfo.InvariantCulture), "Date");
        default: return (value.ToString() ?? string.Empty, "String");
    }
}
```

### B) AutoFit (compute widths in parallel, apply once)

```csharp
public void AutoFitColumns(ExecutionMode? mode = null, CancellationToken ct = default)
{
    var columns = GetAllColumnIndices();
    if (columns.Count == 0) return;

    double[] computed = new double[columns.Count];

    ExecuteWithPolicy(
        opName: "AutoFitColumns",
        itemCount: columns.Count,
        overrideMode: mode,
        sequentialCore: () =>
        {
            for (int i = 0; i < columns.Count; i++)
                computed[i] = CalculateColumnWidthSequential(columns[i]);
            for (int i = 0; i < columns.Count; i++)
                SetColumnWidthCore(columns[i], computed[i]);
            SaveWorksheet();
        },
        computeParallel: () =>
        {
            Parallel.For(0, columns.Count, new ParallelOptions {
                CancellationToken = ct,
                MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
            }, i =>
            {
                computed[i] = CalculateColumnWidthSequential(columns[i]); // compute-only
            });
        },
        applySequential: () =>
        {
            for (int i = 0; i < columns.Count; i++)
                SetColumnWidthCore(columns[i], computed[i]); // DOM write
            SaveWorksheet();
        },
        ct: ct
    );
}
```

### C) Structural ops that stay Sequential

These mutate complex parts and don’t benefit from parallel compute for typical sizes:

* `AddAutoFilter`
* `AddConditionalColorScale`, `AddConditionalDataBar`, `AddConditionalFormatting`
* `Freeze`
* `AddTable`

Keep them using the **sequentialCore** path only (no compute/apply split).

---

## Planner Pattern (SharedStrings / Styles)

**Why:** In parallel compute you must not touch OpenXML parts.
**How:** Build a plan concurrently; apply once.

```csharp
internal sealed class SharedStringPlanner
{
    // Collected during compute
    private readonly ConcurrentDictionary<string, byte> _distinct = new();

    // Final mapping after apply
    private Dictionary<string,int>? _finalIndex;

    public void Note(string s) => _distinct.TryAdd(s, 0);

    public void ApplyTo(ExcelDocument doc)
    {
        // Under lock: fetch existing table, add new entries, build final mapping
        // doc.GetOrCreateSharedStringTable() ...
        // For each s in _distinct.Keys not present, append and capture new index
        // Build _finalIndex mapping for fast lookup
    }

    public void Fixup(ref (int Row, int Col, string Val, string Type) prepared)
    {
        if (prepared.Type != "SharedString") return;
        // Replace text with index text, and optionally set Type="SharedStringIndex"
        var idx = _finalIndex![prepared.Val];
        prepared.Val = idx.ToString(CultureInfo.InvariantCulture);
        // prepared.Type could remain "SharedString" if your core understands index-as-text.
    }

    public void ApplyAndFixup(
        (int Row, int Col, string Val, string Type)[] prepared,
        ExcelSheet sheet)
    {
        // Called inside apply stage, under document lock
        ApplyTo(sheet._document);
        for (int i = 0; i < prepared.Length; i++)
            Fixup(ref prepared[i]);
    }
}
```

Use an analogous planner for **Styles/NumberFormats** if you generate them dynamically.

---

## Fluent API

* Default the **fluent builder** to `ExecutionMode.Sequential` (document-level), so combined operations (AutoFilter + ConditionalFormatting + AutoFit) are stable and lock-free.
* Optionally, make the fluent builder internally use an **internal `BeginNoLock()`** scope to avoid any incidental locking costs.

---

## Implementation Steps

### Phase 1 – Core Infrastructure

1. [ ] Add `ExecutionMode` and `ExecutionPolicy` with `OnDecision`, `MaxDegreeOfParallelism`.
2. [ ] Add `Execution` to `ExcelDocument`; add `ExecutionOverride` to `ExcelSheet`.
3. [ ] Implement `Locking` with `AsyncLocal<bool>` and `ExecuteWrite`.
4. [ ] Implement `ExecuteWithPolicy` (compute/apply split + cancellation).

### Phase 2 – Critical Refactors (Fix current issues)

1. [ ] Refactor `CellValue`/`CellValues` to use the helper and **no nested locks**.
2. [ ] Refactor `AutoFitColumns/Rows` to compute widths in parallel, apply once.
3. [ ] Keep `AddAutoFilter` + Conditional Formatting ops sequential-only.
4. [ ] Verify the original failing fluent example works in `Sequential` mode.

### Phase 3 – Planners

1. [ ] Implement `SharedStringPlanner` and integrate into `CellValues` apply phase.
2. [ ] Implement `StylePlanner` (number formats, fills, fonts) if needed.

### Phase 4 – Batch APIs & Objects

1. [ ] `InsertObjects` uses compute/apply split; parallel flattening; planner usage.
2. [ ] `InsertDataTable` mirrors `CellValues` path, with planner integration.

### Phase 5 – Fluent & Batching

1. [ ] Fluent builder aggregates values and calls `CellValues` in one batch.
2. [ ] Fluent builder defaults to `ExecutionMode.Sequential` (or sets an internal `BeginNoLock()` scope).

### Phase 6 – Polish

1. [ ] Operation defaults for thresholds.
2. [ ] Diagnostics wiring with `OnDecision`.
3. [ ] Benchmarks & unit tests.

---

## Testing Strategy

### Functional

1. **Original failing case**: Fluent with AutoFilter + ConditionalFormatting + AutoFit (Sequential).
2. **Large dataset write**: 200k cells → Automatic (should switch to Parallel).
3. **AutoFit stress**: many columns with long strings → Parallel compute, single apply.
4. **Concurrent writers**: two threads on different sheets using Automatic; ensure no deadlocks.
5. **SharedStrings planner**: repeated strings map to correct indices.

### Concurrency & Safety

* Race tests with random mixes of `CellValues` and structural ops (structural ops forced Sequential).
* Ensure `BeginNoLock()` scope behaves as expected; if kept internal, fluent exercises it.

### Performance

* **Sequential no-lock** vs **Sequential with lock** (legacy) → expect 50–100% faster on small/medium batches.
* **Parallel compute** vs **Sequential** on heavy workloads (AutoFit, object flattening) → expect clear gains with `dop = Environment.ProcessorCount`.

---

## Migration Guide

### Quick Stability Fix

```csharp
using (var document = ExcelDocument.Create(filePath))
{
    // Stability-first default for fluent:
    document.Execution.Mode = ExecutionMode.Sequential;

    document.AsFluent()
        .Sheet("Data", s => s
            .AutoFilter("A1:B3")
            .ConditionalColorScale("B2:B3", Color.Red, Color.Green)
            .AutoFit(columns: true, rows: true))
        .End()
        .Save(openExcel);
}
```

### Performance (Large Batches)

```csharp
// Per-operation thresholds
sheet.ExecutionOverride = new ExecutionPolicy
{
    ParallelThreshold = 10_000,
    OperationThresholds =
    {
        ["AutoFitColumns"] = 2_000,
        ["CellValues"]     = 10_000,
        ["InsertObjects"]  = 1_000
    },
    MaxDegreeOfParallelism = Environment.ProcessorCount
};

// Let Automatic decide:
sheet.CellValues(hugeCells); // compute in parallel, apply once
```

---

## Notes

* OpenXML SDK isn’t thread-safe → **never** mutate DOM from multiple threads.
* Parallelism targets **compute-only** phases; apply is always serialized.
* The planner pattern is essential for SharedStrings/Styles correctness.
* With centralized locking and a compute/apply split, recursion support isn’t needed; it lowers lock overhead.
* Keep public surface small. `CellValue`/`CellValues` (+ `ExecutionMode`) are sufficient. `BeginNoLock()` can be **internal** to your fluent engine.

---

If you want, I can generate a compact benchmark harness (BenchmarkDotNet) tailored to your `CellValues` and `AutoFitColumns` paths so you can capture the before/after deltas immediately.
