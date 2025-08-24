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

1. [x] Add `ExecutionMode` and `ExecutionPolicy` with `OnDecision`, `MaxDegreeOfParallelism`.
2. [x] Add `Execution` to `ExcelDocument`; add `ExecutionOverride` to `ExcelSheet`.
3. [x] Implement `Locking` with `AsyncLocal<bool>` and `ExecuteWrite`.
4. [x] Implement `ExecuteWithPolicy` (compute/apply split + cancellation).

### Phase 2 – Critical Refactors (Fix current issues)

1. [x] Refactor `CellValue`/`CellValues` to use the helper and **no nested locks**.
2. [x] Refactor `AutoFitColumns/Rows` to compute widths in parallel, apply once.
3. [x] Keep `AddAutoFilter` + Conditional Formatting ops sequential-only.
4. [x] Verify the original failing fluent example works in `Sequential` mode.

### Phase 3 – Planners

1. [x] Implement `SharedStringPlanner` and integrate into `CellValues` apply phase.
2. [x] Implement `StylePlanner` (number formats, fills, fonts) if needed.

### Phase 4 – Batch APIs & Objects

1. [x] `InsertObjects` uses compute/apply split (via `SetCellValues`); planner usage. (Flattening currently sequential.)
2. [x] `InsertDataTable` mirrors `CellValues` path, with planner integration.

### Phase 5 – Fluent & Batching

1. [x] Fluent builder aggregates values and calls `CellValues` in one batch.
2. [x] Fluent builder defaults to `ExecutionMode.Sequential` (and sets an internal no‑lock scope).

### Phase 6 – Polish

1. [x] Operation defaults for thresholds.
2. [x] Diagnostics wiring with `OnDecision`.
3. [ ] Benchmarks.
4. [x] Unit tests.

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

---

## Next Steps

- Parallelize object flattening in `InsertObjects` for large inputs (respect `ExecutionPolicy`).
- Add BenchmarkDotNet benchmarks for `CellValues` and `AutoFitColumns/Rows` to quantify gains.
- Expand concurrency tests to mixed-sheet automatic mode decisions and `OnDecision` diagnostics hooks.
- Consider caching/optimizing style application for dynamic number formats under heavy loads.
- Add README examples for the new read APIs (`OfficeIMO.Excel.Read`) and batching patterns.

Status
- Implemented: Execution policy (Automatic/Sequential/Parallel), compute→apply split, SharedStrings/Styles planners, read APIs (range/rows/objects/stream), AutoFit compute parallelization.
- Pending: Multiline AutoFit specifics (wrap and merged spans), reader materializer chain, fluent read surface, benchmarks, and final naming cleanup.

# Quick take (what looks right / what to tweak)

* Your **compute → apply** split is the right direction. Keep all OpenXML DOM mutations in the single serialized “apply” phase only.
* For **reads**, you can be more aggressive: traverse once; offload conversion in chunks; yield lazily. The lazy, chunked, ordered pipeline we sketched fits here.
* **Multiline AutoFit** needs a different algorithm than single-line: you must wrap at the current column width (or merged span) and account for `\n`/`CHAR(10)` + `wrapText`, font metrics, and vertical padding.

# Naming recommendations (tidy & predictable)

## Write API

* BREAKING: Adopt `SetValue(row, col, value)` and `SetValues(batch)` (replace `CellValue/CellValues`).
* Keep `AutoFitColumns`, `AutoFitRows`.
* Structural verbs: `AddTable`, `AddAutoFilter`, `AddConditionalFormatting`, `FreezePanes` (rename from `Freeze`).
* Options bucket stays `ExecutionPolicy` with `Mode/ParallelThreshold/OperationThresholds/MaxDegreeOfParallelism/OnDecision`.

## Read API

Use a **“Read/As/To*”*\* shape that mirrors LINQ/ADO:

* Entry (builder-style chain over a light RangeQuery):

  * `ReadRange("A1:C100")`
  * `ReadRangeStream("A1:C100", chunkRows: 2048)`
* Materializers:

  * `.ToDataTable(headers: true)`
  * `.ToRows()` → `IEnumerable<object?[]>`
  * `.ToObjects()` → `IEnumerable<Dictionary<string,object?>>`
* Low-level:

  * `.AsDataReader()` (new `RangeChunkDataReader`) for `SqlBulkCopy`.

This gives you **discoverable** names and a smooth path from simple reads to ETL.

---

# Fluent Read API (proposal)

```csharp
document.AsFluent()
    .Read(sheet: "Data", r => r
        .Range("A1:Z1000000")
        .Headers(true)
        .ChunkRows(4096)
        .Execution(ExecutionMode.Automatic)
        .Into(dt => /* got DataTable */)
        // or:
        //.IntoRows(rows => ...)
        //.IntoObjects(dicts => ...)
        //.IntoDataReader(dr => bulkCopy.WriteToServer(dr))
    )
    .End();
```

Internally this just wraps your `ExcelSheetReader.ReadRange` / `ReadRangeStream` + materializers.

---

# AutoFit that handles multiline (what Excel actually does)

You need **two pieces**:

1. **Width** (columns): compute the maximum single-line pixel width among cells **formatted with that column’s style**, then convert to Excel column width units (characters of “0” in Normal style) with the standard fudge factors. Excel’s spec is quirky; see Eric White’s write-up and OOXML `col` width details. ([ericwhite.com][1], [Microsoft Learn][2], [Stack Overflow][3])

2. **Height** (rows) with wrapping:

   * For each row, for each cell:

     * Determine the **effective wrapping width** in pixels:

       * If **merged** horizontally, width = sum of spanned column widths + inter-cell gridlines.
       * Else width = that column’s pixel width.
     * If the cell has **manual line breaks** (`\n` / `CHAR(10)`), split first.
     * For each line fragment, perform **soft wrapping** at the measured width using word-boundary rules.
     * Count resulting wrapped lines → `lineCount`.
     * Compute height = `topPadding + lineCount * lineHeight + bottomPadding` (line height from font metrics + leading).
   * Row height = **max** of all cell heights in the row (unless locked by an explicit height).
   * Respect Excel quirks:

     * `wrapText=false` → single-line (truncate visually, height unchanged).
     * `ShrinkToFit` reduces font rendering width, but Excel’s logic is non-trivial; for now, ignore or handle later.
     * Merged cells + wrap: Excel has edge cases (Excel won’t auto-fit height when the cell is merged in certain scenarios). Document this caveat.

**Implementation choices for measurement:**

* Default: SixLabors.Fonts-based measurer (already used in repo) for cross-platform text metrics.
* Optional: **SkiaSharp** (`SKPaint.MeasureText`, `BreakText`) to stay deterministic.
* Fallback: heuristic approximation (EPPlus-style) if neither engine is available. ([GitHub][4])

**Column width conversions and constants**: Excel uses a base of the **max digit** glyph width of Normal style (Calibri 11 → \~7 px at 96 DPI) and `Truncate(chars * maxDigit) + padding`. Keep this centralized. ([Microsoft Learn][5])

Microsoft/OOXML references on widths/heights: ([Eric White][1], [Microsoft Learn][2], [Stack Overflow][6])

---

# Parallel model sanity-check (reads & writes)

**Writes**

* ✅ Parallel **compute only** (coercion, CF evaluation, AutoFit measurement).
* ✅ Single **apply** phase under one coarse lock.
* ✅ SharedStrings/Styles use a **planner** (distinct gather → apply once → fixup indexes).
* ⚠️ Explicit rule: no OpenXML part mutation during compute (e.g., `GetSharedStringIndex`, styles). Use planners to gather, then mutate once in apply.

**Reads**

* ✅ DOM traversal single-threaded → chunk raw.
* ✅ Chunk conversion offloaded with **bounded parallelism**.
* ✅ **Ordered** delivery (chunk index) for stable consumers.
* ✅ `CancellationToken` + `MaxDegreeOfParallelism`.

If any part of write-compute touches `WorksheetPart`, `SharedStringTablePart`, or `WorkbookStylesPart`, move it into the apply stage.

---

# Targeted improvements to drop in now

1. **Rename write surface (breaking)**

   * `CellValue` → `SetValue`
   * `CellValues` → `SetValues`
   * Update fluent and tests accordingly (no shims necessary).

2. **Reader materializers**

   * Add `ReadRange(...).ToDataTable(headers: bool)`,
     `ReadRange(...).ToRows()`, `ReadRange(...).ToObjects()`,
     and `ReadRangeStream(...).AsDataReader()` using a light `RangeQuery/RangeStreamQuery` wrapper.

3. **AutoFit v2**

   * Introduce `ITextMeasurer` with a default **SixLabors** implementation; optional Skia engine.
   * Add `AutoFitOptions`:

     * `MeasureEngine: "Skia"|"GDI"|"Heuristic"`
     * `IncludeMergedCells: bool`
     * `WrapAtMergedSpan: bool`
     * `RespectShrinkToFit: bool` (future)
   * New internals:

     * `MeasureSingleLineWidth(text, font)` → px
     * `WrapLines(text, font, widthPx)` → `IReadOnlyList<string>`
     * `ComputeRowHeight(cell, colSpanWidthPx, style)` → px
     * Conversion helpers `Pixels↔ExcelUnits` centralized.

4. **Policy defaults**

   * `Execution.Policy.OperationThresholds["CellValues"] = 10_000;`
   * `["AutoFitColumns"] = 2_000; ["AutoFitRows"] = 2_000;`
   * `["InsertObjects"] = 1_000;`
   * `MaxDegreeOfParallelism = Environment.ProcessorCount`.

5. **Benchmarks**

   * Bench harness for:

     * `SetValues` small vs large
     * `AutoFitColumns/Rows` with/without wrap
     * `ReadRange` vs `ReadRangeStream` (parallel).

---

# New TODO (laser-focused)

### A. Naming & API surface (breaking)

* [ ] Rename `CellValue/CellValues` → `SetValue/SetValues` across code/tests/fluent.
* [ ] Structural names unified: `AddTable`, `AddAutoFilter`, `AddConditionalFormatting`, `FreezePanes` (rename from `Freeze`).
* [ ] Reader materializers: `ToDataTable`, `ToRows`, `ToObjects`, `AsDataReader` via `RangeQuery` wrappers.

### B. Fluent Read

* [ ] `Fluent.Read(...)` with `.Range()`, `.Headers()`, `.ChunkRows()`, `.Execution()`, `.Into(...)`.
* [ ] Expose `AsDataReader` path in fluent for bulk copy.

### C. Parallel correctness

* [ ] Audit: **no** OpenXML part mutation outside apply stage.
* [ ] SharedStrings/Styles **planner** integrated in apply stage only.
* [ ] Reads: ensure `OpenXmlReader` or DOM traversal is **single-threaded**; conversion is chunk-parallel.

### D. AutoFit v2 (multiline, merged cells)

* [ ] Introduce `ITextMeasurer` + `SixLaborsTextMeasurer` (and optional `SkiaTextMeasurer`).
* [ ] Implement word-wrap + manual breaks (`\n`/`CHAR(10)`) respecting `wrapText`.
* [ ] Account for **merged spans** when computing wrap width.
* [ ] Compute row height as max of wrapped cells; convert px→row height.
* [ ] Centralize conversions (`ExcelColWidth↔px`, `RowHeight↔px`) with constants/Normal style.
* [ ] Add `AutoFitOptions` (engine, merged handling, future `ShrinkToFit`).

### E. Heuristics fallback (optional)

* [ ] Implement a **heuristic** measurer (no native deps) to keep pure-managed option.

---

References
1. Eric White on column widths and Excel units: https://ericwhite.com/blog/ and archived OpenXML sizing posts
2. Microsoft Learn – Change column width: https://learn.microsoft.com/office/open-xml/how-to-change-the-width-of-a-column-in-a-spreadsheet
3. SkiaSharp MeasureText: https://learn.microsoft.com/dotnet/api/skiasharp.skpaint.measuretext
4. EPPlus width/height discussions (heuristics): https://github.com/EPPlusSoftware/EPPlus
5. Excel column width and row height overview: https://support.microsoft.com/office/change-the-column-width-and-row-height-c2c3f030-5ab9-4d2a-8dfd-f5b6a25ac749
6. Stack Overflow – OpenXML column width units: https://stackoverflow.com/questions/17888225/column-width-in-openxml

### F. Diagnostics & perf

* [ ] `ExecutionPolicy.OnDecision` hooks → log op name, count, mode.
* [ ] BenchmarkDotNet: publish `Before/After` for:

  * `SetValues` small (≤1k) vs big (≥100k)
  * `AutoFitRows` on wrapped text
  * `ReadRange` vs `ReadRangeStream`.

### G. Tests

* [ ] Row height correctness with:

  * Single line, wrap off
  * Multiline via `\n`
  * Soft wrap at width (long words, long URLs)
  * Merged cells (2–5 columns)
* [ ] Column width correctness on mixed fonts/styles
* [ ] Parallel reads/writes correctness under cancellation.

---

# Small code sketch (interfaces you can paste)

```csharp
public interface ITextMeasurer
{
    TextMetrics MeasureSingleLine(string text, FontSpec font);
    WrappedText Wrap(string text, double availableWidthPx, FontSpec font);
}

public sealed record FontSpec(string Family, double SizePt, bool Bold, bool Italic);
public sealed record TextMetrics(double WidthPx, double LineHeightPx);
public sealed record WrappedText(IReadOnlyList<string> Lines, TextMetrics Metrics);

public sealed class AutoFitOptions
{
    public string MeasureEngine { get; init; } = "Skia"; // or "GDI", "Heuristic"
    public bool IncludeMergedCells { get; init; } = true;
    public bool WrapAtMergedSpan { get; init; } = true;
    public bool RespectShrinkToFit { get; init; } = false; // future
}

public static class ExcelUnits
{
    // Centralize conversions; constants taken from OOXML/Excel behavior.
    public static double ColumnWidthCharsToPixels(double chars, double maxDigitPx = 7.0, double paddingPx = 5.0)
        => Math.Floor(chars * maxDigitPx) + paddingPx;
    public static double PixelsToColumnWidthChars(double px, double maxDigitPx = 7.0, double paddingPx = 5.0)
        => (px - paddingPx) / maxDigitPx;

    // Row height is stored in points; 1pt = 96/72 px at 100% scaling
    public static double PointsToPixels(double pt) => pt * (96.0 / 72.0);
    public static double PixelsToPoints(double px) => px * (72.0 / 96.0);
}
```

Then in `AutoFitRows`:

```csharp
public void AutoFitRows(AutoFitOptions? options = null, ExecutionMode? mode = null)
{
    options ??= new AutoFitOptions();
    var measurer = TextMeasurerFactory.Create(options.MeasureEngine);

    // compute in parallel: gather effective widths per cell (consider merged spans),
    // wrap text, compute required height per cell, reduce to max per row
    // apply once: set row heights (points) + save worksheet
}
```

References for width/height math and OOXML `col` behavior: Eric White’s write-up and the OOXML column spec; both are the best grounding for the constants and conversions. ([ericwhite.com][1], [Microsoft Learn][2])
Some practical discussions/heuristics also exist (StackOverflow/answers). ([Stack Overflow][6], [Microsoft Learn][5])


[1]: https://www.ericwhite.com/blog/precisely-calculating-cell-sizes-in-open-xml-spreadsheetml/?utm_source=chatgpt.com "Precisely Calculating Cell Sizes in Open XML ..."
[2]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1&utm_source=chatgpt.com "Column Class (DocumentFormat.OpenXml.Spreadsheet)"
[3]: https://stackoverflow.com/questions/7716078/formula-to-convert-net-pixels-to-excel-width-in-openxml-format/53741810?utm_source=chatgpt.com "Formula to convert .NET pixels to Excel width in OpenXML ..."
[4]: https://github.com/EPPlusSoftware/EPPlus/wiki/Autofit-columns?utm_source=chatgpt.com "Autofit columns · EPPlusSoftware/EPPlus Wiki"
[5]: https://learn.microsoft.com/en-us/answers/questions/4973888/%28excel%29-how-to-calculate-the-default-column-width?utm_source=chatgpt.com "(Excel) How to calculate the default column width for a ..."
[6]: https://stackoverflow.com/questions/34374785/creating-custom-column-widths-in-openxml-excel?utm_source=chatgpt.com "Creating custom column widths in OpenXML (excel)"
