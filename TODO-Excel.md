# TODO: Excel Processing Refactor - Parallel vs Sequential Control

## Problem Summary
The Excel fluent API fails to open files when certain operations are combined (AutoFilter, ConditionalFormatting, AutoFit, Column operations). Investigation revealed:
- All operations (parallel and sequential) use the same `ReaderWriterLockSlim` instance
- Locks are acquired even for non-concurrent operations (unnecessary overhead)
- The parallel threshold (1000 cells) is rarely reached in typical fluent API usage
- Locking issues affect both parallel and sequential code paths equally
- Nested lock acquisitions with inconsistent recursion handling cause deadlocks

## Root Cause
1. **Unnecessary locking**: Sequential operations don't need thread synchronization
2. **Shared lock instance**: Both parallel and sequential paths use the same `_lock`
3. **Inconsistent lock handling**: Some methods check `IsWriteLockHeld`, others don't
4. **Multiple structural modifications**: Operations like AutoFilter + AutoFit + ConditionalFormatting conflict

## Architecture Design

### Core Components

#### 1. Execution Policy
```csharp
public enum ExecutionMode
{
    Automatic,   // Decide by thresholds (default)
    Sequential,  // Force single-threaded, no locks
    Parallel     // Force parallel with locks
}

public sealed class ExecutionPolicy
{
    public ExecutionMode Mode { get; set; } = ExecutionMode.Automatic;
    
    /// <summary>
    /// Default threshold for cells/items above which Automatic switches to Parallel.
    /// </summary>
    public int ParallelThreshold { get; set; } = 1000;
    
    /// <summary>
    /// Optional per-operation overrides (e.g., "CellValues", "InsertObjects", "AutoFitColumns").
    /// </summary>
    public Dictionary<string, int> OperationThresholds { get; } = new(StringComparer.Ordinal);
    
    /// <summary>
    /// Enable diagnostics to log execution decisions.
    /// </summary>
    public bool EnableDiagnostics { get; set; }
    
    internal void LogDecision(string operation, int itemCount, ExecutionMode decided)
    {
        if (EnableDiagnostics)
        {
            Debug.WriteLine($"[Excel] {operation}: {itemCount} items → {decided}");
        }
    }
}
```

#### 2. Document and Sheet Integration
```csharp
public class ExcelDocument : IDisposable
{
    /// <summary>
    /// Global execution policy for all sheets.
    /// </summary>
    public ExecutionPolicy Execution { get; } = new();
    
    /// <summary>
    /// Single lock for cross-sheet structural operations (allocated lazily).
    /// </summary>
    internal ReaderWriterLockSlim? _lock;
    
    internal ReaderWriterLockSlim EnsureLock()
    {
        return _lock ??= new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
    }
}

public partial class ExcelSheet
{
    /// <summary>
    /// Sheet-specific execution policy override.
    /// null = inherit from document.
    /// </summary>
    public ExecutionPolicy? ExecutionOverride { get; set; }
    
    internal ExecutionPolicy EffectiveExecution => ExecutionOverride ?? _document.Execution;
}
```

#### 3. Locking Infrastructure
```csharp
internal static class Locking
{
    // AsyncLocal works across async/await without leaking between threads
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
    
    public static void ExecuteWrite(ReaderWriterLockSlim? lck, Action action)
    {
        // Skip locking if in NoLock scope or no lock allocated
        if (IsNoLock || lck is null)
        {
            action();
            return;
        }
        
        // Skip if already holding write lock (recursive call)
        if (lck.IsWriteLockHeld)
        {
            action();
            return;
        }
        
        // Acquire lock for this operation
        lck.EnterWriteLock();
        try { action(); }
        finally { lck.ExitWriteLock(); }
    }
}
```

#### 4. NoLock Context for Power Users
```csharp
public partial class ExcelSheet
{
    /// <summary>
    /// Begin a no-lock scope where all operations skip synchronization.
    /// User is responsible for thread safety.
    /// </summary>
    public NoLockContext BeginNoLock() => new();
    
    public sealed class NoLockContext : IDisposable
    {
        private readonly IDisposable _scope;
        internal NoLockContext() => _scope = Locking.EnterNoLockScope();
        public void Dispose() => _scope.Dispose();
    }
}

// Usage:
using (sheet.BeginNoLock())
{
    // All operations here skip locks, even nested calls
    for (int i = 1; i <= 10000; i++)
        sheet.CellValueNoLock(i, 1, $"Value {i}");
    
    sheet.AutoFitColumnsNoLock();
    sheet.AddTableNoLock(...);
}
```

### Method Implementation Pattern

#### The Three-Layer Pattern
Every operation follows this consistent pattern:

```csharp
public partial class ExcelSheet
{
    // 1. SAFE PUBLIC API (with locks)
    public void MethodName(params..., ExecutionMode? mode = null)
    {
        ExecuteBatchOrWriteLocked(
            workSequentialCore: _ => MethodNameCore(params...),
            itemCountHint: EstimateItemCount(),
            opName: "MethodName",
            overrideMode: mode,
            workParallelCore: _ => MethodNameParallelCore(params...)
        );
    }
    
    // 2. NO-LOCK PUBLIC API (for power users)
    public void MethodNameNoLock(params...)
    {
        MethodNameCore(params...);
    }
    
    // 3. CORE IMPLEMENTATION (no locks, actual work)
    private void MethodNameCore(params...)
    {
        // Direct DOM manipulation
        // This is the single source of truth
    }
    
    // 4. Optional: Parallel-specific implementation
    private void MethodNameParallelCore(params...)
    {
        // Compute in parallel (no DOM mutations)
        var results = ComputeInParallel(params...);
        
        // Apply results sequentially (DOM not thread-safe)
        foreach (var result in results)
            ApplyResult(result);
    }
}
```

#### Batch Execution Helper
```csharp
private void ExecuteBatchOrWriteLocked(
    Action<ExecutionPolicy> workSequentialCore,
    int itemCountHint,
    string opName,
    ExecutionMode? overrideMode = null,
    Action<ExecutionPolicy>? workParallelCore = null)
{
    var policy = EffectiveExecution;
    var mode = overrideMode ?? policy.Mode;
    
    // Auto-decide based on threshold
    if (mode == ExecutionMode.Automatic)
    {
        var threshold = policy.OperationThresholds.TryGetValue(opName, out var v) 
            ? v 
            : policy.ParallelThreshold;
        mode = itemCountHint > threshold 
            ? ExecutionMode.Parallel 
            : ExecutionMode.Sequential;
        
        policy.LogDecision(opName, itemCountHint, mode);
    }
    
    if (mode == ExecutionMode.Sequential)
    {
        // Run without any locks
        using (Locking.EnterNoLockScope())
            workSequentialCore(policy);
        return;
    }
    
    // Parallel: acquire ONE lock for entire operation
    Locking.ExecuteWrite(_document.EnsureLock(), () => 
    {
        (workParallelCore ?? workSequentialCore)(policy);
    });
}
```

### Concrete Implementation Examples

#### Cell Operations
```csharp
// SAFE API
public void CellValue(int row, int column, object value, ExecutionMode? mode = null)
{
    ExecuteBatchOrWriteLocked(
        workSequentialCore: _ => CellValueCore(row, column, value),
        itemCountHint: 1,
        opName: "CellValue",
        overrideMode: mode
    );
}

// NO-LOCK API
public void CellValueNoLock(int row, int column, object value)
    => CellValueCore(row, column, value);

// CORE
private void CellValueCore(int row, int column, object value)
{
    Cell cell = GetCell(row, column);
    switch (value)
    {
        case string s:
            int sharedStringIndex = _excelDocument.GetSharedStringIndex(s);
            cell.CellValue = new CellValue(sharedStringIndex.ToString());
            cell.DataType = CellValues.SharedString;
            break;
        case double d:
            cell.CellValue = new CellValue(d.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
            break;
        // ... other types
    }
}

// BATCH API
public void CellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null)
{
    var list = cells as IList<(int,int,object)> ?? cells.ToList();
    
    ExecuteBatchOrWriteLocked(
        workSequentialCore: _ =>
        {
            foreach (var (r, c, v) in list)
                CellValueCore(r, c, v);
        },
        itemCountHint: list.Count,
        opName: "CellValues",
        overrideMode: mode,
        workParallelCore: _ =>
        {
            // Prepare values in parallel (compute only)
            var prepared = new ConcurrentBag<CellUpdate>();
            Parallel.ForEach(list, item =>
            {
                prepared.Add(PrepareCellUpdate(item.Row, item.Column, item.Value));
            });
            
            // Apply to DOM sequentially (under lock)
            foreach (var update in prepared)
                ApplyCellUpdate(update);
        }
    );
}
```

#### Structural Operations
```csharp
// SAFE API
public void AutoFitColumns(ExecutionMode? mode = null)
{
    ExecuteBatchOrWriteLocked(
        workSequentialCore: _ => AutoFitColumnsCore(),
        itemCountHint: EstimateColumnCount(),
        opName: "AutoFitColumns",
        overrideMode: mode,
        workParallelCore: _ =>
        {
            // Compute widths in parallel (no DOM mutation)
            var widths = ComputeColumnWidthsParallel();
            
            // Apply widths sequentially (DOM mutation)
            foreach (var (columnIndex, width) in widths)
                SetColumnWidthCore(columnIndex, width);
        }
    );
}

// NO-LOCK API
public void AutoFitColumnsNoLock() => AutoFitColumnsCore();

// CORE
private void AutoFitColumnsCore()
{
    var worksheet = _worksheetPart.Worksheet;
    var sheetData = worksheet.GetFirstChild<SheetData>();
    if (sheetData == null) return;
    
    // Get all column indices
    var columnIndices = GetAllColumnIndices(sheetData);
    
    // Calculate and apply width for each column
    foreach (int index in columnIndices)
    {
        double width = CalculateColumnWidth(index);
        SetColumnWidthCore(index, width);
    }
    
    worksheet.Save();
}

private void SetColumnWidthCore(int columnIndex, double width)
{
    // Direct DOM manipulation, no locks
    var worksheet = _worksheetPart.Worksheet;
    var columns = worksheet.GetFirstChild<Columns>() 
        ?? worksheet.InsertAt(new Columns(), 0);
    
    // ... update column width
}
```

### Object/Data Insertion Enhancement

```csharp
public class ObjectInsertionOptions
{
    public bool IncludeHeaders { get; set; } = true;
    public int StartRow { get; set; } = 1;
    public int StartColumn { get; set; } = 1;
    public ExecutionMode? Mode { get; set; }  // null = auto-decide
    public bool ExpandNestedObjects { get; set; } = true;
    public int MaxNestingDepth { get; set; } = 3;
    public List<string>? IncludeProperties { get; set; }  // Whitelist
    public List<string>? ExcludeProperties { get; set; }  // Blacklist
    public Func<string, string>? HeaderTransform { get; set; }  // Custom headers
    public bool CreateTable { get; set; } = false;
    public TableStyle TableStyle { get; set; } = TableStyle.Medium2;
}

// SAFE API
public void InsertObjects<T>(IEnumerable<T> items, ObjectInsertionOptions? options = null)
{
    options ??= new ObjectInsertionOptions();
    var cells = PrepareObjectCells(items, options);
    
    ExecuteBatchOrWriteLocked(
        workSequentialCore: _ =>
        {
            foreach (var cell in cells)
                CellValueCore(cell.Row, cell.Column, cell.Value);
            
            if (options.CreateTable)
                AddTableCore(/*...*/);
        },
        itemCountHint: cells.Count,
        opName: "InsertObjects",
        overrideMode: options.Mode,
        workParallelCore: _ =>
        {
            // Prepare in parallel
            var prepared = PrepareObjectCellsParallel(items, options);
            
            // Apply sequentially
            foreach (var cell in prepared)
                CellValueCore(cell.Row, cell.Column, cell.Value);
            
            if (options.CreateTable)
                AddTableCore(/*...*/);
        }
    );
}

// NO-LOCK API
public void InsertObjectsNoLock<T>(IEnumerable<T> items, ObjectInsertionOptions? options = null)
{
    options ??= new ObjectInsertionOptions();
    var cells = PrepareObjectCells(items, options);
    
    foreach (var cell in cells)
        CellValueCore(cell.Row, cell.Column, cell.Value);
    
    if (options.CreateTable)
        AddTableCore(/*...*/);
}
```

## Methods Requiring Refactoring

All these methods need the three-layer pattern (Safe/NoLock/Core):

### Cell Operations
- [x] `CellValue` (all overloads)
- [x] `CellValues` (batch)
- [ ] `CellFormula`
- [ ] `FormatCell`

### Column/Row Operations
- [x] `AutoFitColumns`
- [ ] `AutoFitRows`
- [ ] `SetColumnWidth`
- [ ] `SetColumnHidden`
- [ ] `SetRowHeight`
- [ ] `AutoFitColumn` (single)
- [ ] `AutoFitRow` (single)

### Formatting Operations
- [ ] `Freeze`
- [ ] `AddAutoFilter`
- [ ] `AddConditionalColorScale`
- [ ] `AddConditionalDataBar`
- [ ] `AddConditionalFormatting`

### Table Operations
- [ ] `AddTable`
- [x] `InsertObjects`
- [ ] `InsertDataTable`

### Fluent API
- [ ] Update `SheetBuilder` to use NoLock internally
- [ ] Add `ExecutionMode` parameter to fluent methods
- [ ] Consider batch collection in fluent context

## Implementation Steps

### Phase 1: Core Infrastructure (Non-Breaking)
1. [ ] Add `ExecutionMode` enum
2. [ ] Add `ExecutionPolicy` class
3. [ ] Add `Execution` property to `ExcelDocument`
4. [ ] Add `ExecutionOverride` property to `ExcelSheet`
5. [ ] Implement `Locking` helper class with `AsyncLocal`
6. [ ] Implement `NoLockContext` and `BeginNoLock()`
7. [ ] Implement `ExecuteBatchOrWriteLocked` helper

### Phase 2: Refactor Critical Methods (Fix Issues)
1. [ ] Refactor `CellValue` methods to Core pattern
2. [ ] Refactor `AutoFitColumns/Rows` to fix nested locking
3. [ ] Refactor `AddAutoFilter` to Core pattern
4. [ ] Refactor `AddConditionalColorScale/DataBar` to Core pattern
5. [ ] Test with problematic fluent example

### Phase 3: Add NoLock APIs (Performance)
1. [ ] Add `*NoLock` methods for all refactored operations
2. [ ] Document NoLock usage and safety requirements
3. [ ] Add examples showing performance benefits

### Phase 4: Enhance Object Insertion
1. [ ] Implement `ObjectInsertionOptions`
2. [ ] Add property filtering and transformation
3. [ ] Add automatic table creation
4. [ ] Support hashtables and dictionaries

### Phase 5: Complete Refactoring
1. [ ] Refactor remaining methods to Core pattern
2. [ ] Add ExecutionMode parameter to batch methods
3. [ ] Update fluent API to use NoLock internally
4. [ ] Add operation-specific thresholds

### Phase 6: Polish and Documentation
1. [ ] Add diagnostic logging
2. [ ] Create performance benchmarks
3. [ ] Write migration guide
4. [ ] Add unit tests for all patterns

## Testing Strategy

### Test Cases
1. **Original failing case**: Fluent API with AutoFilter + ConditionalFormatting + AutoFit
2. **Sequential mode**: Same operations with `ExecutionMode.Sequential`
3. **NoLock context**: Using `BeginNoLock()` with multiple operations
4. **Large dataset**: > 10,000 cells with Parallel vs Sequential
5. **Concurrent access**: Multiple threads with proper locking
6. **Nested operations**: AutoFit inside batch operations
7. **Object insertion**: Complex nested objects with filtering

### Performance Benchmarks
```csharp
// Benchmark: Sequential with locks vs without locks
[Benchmark]
public void CellValues_Sequential_WithLocks()
{
    sheet.Execution.Mode = ExecutionMode.Parallel; // Forces locks
    sheet.CellValues(cells, ExecutionMode.Sequential);
}

[Benchmark]
public void CellValues_Sequential_NoLocks()
{
    using (sheet.BeginNoLock())
    {
        foreach (var cell in cells)
            sheet.CellValueNoLock(cell.Row, cell.Column, cell.Value);
    }
}

// Expected: NoLocks should be 50-100% faster for small batches
```

## Migration Guide

### Quick Fix for Current Issues
```csharp
// If experiencing issues with Excel not opening:
using (var document = ExcelDocument.Create(filePath))
{
    // Option 1: Disable parallel globally
    document.Execution.Mode = ExecutionMode.Sequential;
    
    // Option 2: Use NoLock context
    using (sheet.BeginNoLock())
    {
        // Your fluent API code here - no locks!
        document.AsFluent()
            .Sheet("Data", s => s
                .AutoFilter("A1:B3")
                .ConditionalColorScale("B2:B3", Color.Red, Color.Green)
                .AutoFit(columns: true, rows: true)
            )
            .End();
    }
    
    document.Save(openExcel);
}
```

### Performance Optimization
```csharp
// Configure per-operation thresholds
sheet.ExecutionOverride = new ExecutionPolicy
{
    ParallelThreshold = 500,
    OperationThresholds = 
    {
        ["AutoFitColumns"] = 20,    // Benefits from parallel earlier
        ["CellValues"] = 1000,       // Needs more to justify parallel
        ["InsertObjects"] = 100      // Object processing benefits from parallel
    }
};

// Force sequential for small operations
sheet.CellValues(smallData, ExecutionMode.Sequential);

// Force parallel for large operations
sheet.CellValues(largeData, ExecutionMode.Parallel);

// Let it auto-decide (default)
sheet.CellValues(data);
```

### Power User Optimization
```csharp
// Maximum performance for bulk operations
using (sheet.BeginNoLock())
{
    // Insert million cells with zero lock overhead
    for (int row = 1; row <= 1000; row++)
    {
        for (int col = 1; col <= 1000; col++)
        {
            sheet.CellValueNoLock(row, col, $"R{row}C{col}");
        }
    }
    
    // Structural operations also skip locks
    sheet.AutoFitColumnsNoLock();
    sheet.AddTableNoLock("A1:ALL1000", true, "DataTable", TableStyle.Medium2);
}
```

## Examples Update Guide

### Example Structure
All examples should demonstrate the three usage patterns:

```csharp
namespace OfficeIMO.Examples.Excel {
    internal static class ExampleClassName {
        /// <summary>
        /// Basic example - uses default safe API (backward compatible)
        /// </summary>
        public static void Example_Basic(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Basic example with automatic execution");
            string filePath = Path.Combine(folderPath, "Basic.xlsx");
            
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                
                // Default usage - automatic decision
                sheet.CellValue(1, 1, "Header");
                sheet.CellValues(data);  // Auto-decides based on count
                sheet.AutoFitColumns();   // Safe with locks
                
                document.Save(openExcel);
            }
        }
        
        /// <summary>
        /// Performance example - demonstrates explicit control
        /// </summary>
        public static void Example_Performance(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Performance example with explicit control");
            string filePath = Path.Combine(folderPath, "Performance.xlsx");
            
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                // Configure execution policy
                document.Execution.Mode = ExecutionMode.Sequential;  // No locks for entire document
                // OR configure per-operation thresholds
                document.Execution.OperationThresholds["AutoFitColumns"] = 10;
                
                var sheet = document.AddWorkSheet("Data");
                
                // Explicit control per operation
                sheet.CellValues(smallData, ExecutionMode.Sequential);  // Force sequential
                sheet.CellValues(largeData, ExecutionMode.Parallel);    // Force parallel
                
                document.Save(openExcel);
            }
        }
        
        /// <summary>
        /// Power user example - maximum performance with NoLock
        /// </summary>
        public static void Example_NoLock(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - NoLock example for maximum performance");
            string filePath = Path.Combine(folderPath, "NoLock.xlsx");
            
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                
                // Maximum performance - no locks at all
                using (sheet.BeginNoLock()) {
                    // Bulk operations with zero lock overhead
                    for (int i = 1; i <= 10000; i++) {
                        sheet.CellValueNoLock(i, 1, $"Row {i}");
                    }
                    
                    sheet.AutoFitColumnsNoLock();
                    sheet.AddTableNoLock("A1:A10000", true, "DataTable");
                }
                
                document.Save(openExcel);
            }
        }
    }
}
```

### Fluent API Examples

```csharp
internal static partial class FluentWorkbook {
    /// <summary>
    /// Original example - works but may have locking issues
    /// </summary>
    public static void Example_FluentWorkbook_Legacy(string folderPath, bool openExcel) {
        Console.WriteLine("[*] Excel - Creating workbook with fluent API (legacy)");
        string filePath = Path.Combine(folderPath, "FluentWorkbook_Legacy.xlsx");
        
        using (ExcelDocument document = ExcelDocument.Create(filePath)) {
            document.AsFluent()
                .Sheet("Data", s => s
                    .HeaderRow("Name", "Score")
                    .Row(r => r.Values("Alice", 93))
                    .Row(r => r.Values("Bob", 88))
                    .Table(t => t.Add("A1:B3", true, "Scores"))
                    .Freeze(topRows: 1, leftCols: 1)
                    // These might cause issues due to locking:
                    .AutoFilter("A1:B3")
                    .ConditionalColorScale("B2:B3", Color.Red, Color.Green)
                    .ConditionalDataBar("B2:B3", Color.Blue)
                    .AutoFit(columns: true, rows: true)
                    .Columns(c => c.Col(1, col => col.Width(25)).Col(2, col => col.Hidden(true)))
                )
                .End()
                .Save(openExcel);
        }
    }
    
    /// <summary>
    /// Fixed example - using Sequential mode to avoid locking issues
    /// </summary>
    public static void Example_FluentWorkbook_Fixed(string folderPath, bool openExcel) {
        Console.WriteLine("[*] Excel - Creating workbook with fluent API (fixed)");
        string filePath = Path.Combine(folderPath, "FluentWorkbook_Fixed.xlsx");
        
        using (ExcelDocument document = ExcelDocument.Create(filePath)) {
            // FIX: Set sequential mode to avoid locking issues
            document.Execution.Mode = ExecutionMode.Sequential;
            
            document.AsFluent()
                .Sheet("Data", s => s
                    .HeaderRow("Name", "Score")
                    .Row(r => r.Values("Alice", 93))
                    .Row(r => r.Values("Bob", 88))
                    .Table(t => t.Add("A1:B3", true, "Scores"))
                    .Freeze(topRows: 1, leftCols: 1)
                    // Now these work without issues:
                    .AutoFilter("A1:B3")
                    .ConditionalColorScale("B2:B3", Color.Red, Color.Green)
                    .ConditionalDataBar("B2:B3", Color.Blue)
                    .AutoFit(columns: true, rows: true)
                    .Columns(c => c.Col(1, col => col.Width(25)).Col(2, col => col.Hidden(true)))
                )
                .End()
                .Save(openExcel);
        }
    }
    
    /// <summary>
    /// Optimized example - using NoLock context for fluent API
    /// </summary>
    public static void Example_FluentWorkbook_Optimized(string folderPath, bool openExcel) {
        Console.WriteLine("[*] Excel - Creating workbook with fluent API (optimized)");
        string filePath = Path.Combine(folderPath, "FluentWorkbook_Optimized.xlsx");
        
        using (ExcelDocument document = ExcelDocument.Create(filePath)) {
            var sheet = document.AddWorkSheet("Data");
            
            // Maximum performance with NoLock
            using (sheet.BeginNoLock()) {
                // Future: Fluent API could internally use NoLock
                document.AsFluent()
                    .Sheet(sheet, s => s
                        .HeaderRow("Name", "Score")
                        .Row(r => r.Values("Alice", 93))
                        .Row(r => r.Values("Bob", 88))
                        // All operations run without locks
                    )
                    .End();
            }
            
            document.Save(openExcel);
        }
    }
}
```

### Performance Comparison Examples

```csharp
internal static class PerformanceComparison {
    public static void Example_CompareExecutionModes(string folderPath) {
        Console.WriteLine("[*] Excel - Performance comparison of execution modes");
        
        var data = GenerateLargeDataset(10000);
        var sw = new Stopwatch();
        
        // Test 1: Default (with locks)
        using (var doc = ExcelDocument.Create(Path.Combine(folderPath, "Default.xlsx"))) {
            var sheet = doc.AddWorkSheet("Data");
            
            sw.Restart();
            foreach (var item in data) {
                sheet.CellValue(item.Row, item.Col, item.Value);  // 10,000 lock acquisitions
            }
            sw.Stop();
            Console.WriteLine($"  Default (with locks): {sw.ElapsedMilliseconds}ms");
            
            doc.Save();
        }
        
        // Test 2: Sequential mode (no locks)
        using (var doc = ExcelDocument.Create(Path.Combine(folderPath, "Sequential.xlsx"))) {
            doc.Execution.Mode = ExecutionMode.Sequential;
            var sheet = doc.AddWorkSheet("Data");
            
            sw.Restart();
            foreach (var item in data) {
                sheet.CellValue(item.Row, item.Col, item.Value);  // 0 lock acquisitions
            }
            sw.Stop();
            Console.WriteLine($"  Sequential (no locks): {sw.ElapsedMilliseconds}ms");
            
            doc.Save();
        }
        
        // Test 3: Batch with auto-decision
        using (var doc = ExcelDocument.Create(Path.Combine(folderPath, "Batch.xlsx"))) {
            var sheet = doc.AddWorkSheet("Data");
            
            sw.Restart();
            sheet.CellValues(data);  // Decides based on count (10,000 > 1000 = parallel)
            sw.Stop();
            Console.WriteLine($"  Batch (auto-decide): {sw.ElapsedMilliseconds}ms");
            
            doc.Save();
        }
        
        // Test 4: NoLock maximum performance
        using (var doc = ExcelDocument.Create(Path.Combine(folderPath, "NoLock.xlsx"))) {
            var sheet = doc.AddWorkSheet("Data");
            
            sw.Restart();
            using (sheet.BeginNoLock()) {
                foreach (var item in data) {
                    sheet.CellValueNoLock(item.Row, item.Col, item.Value);  // Direct DOM access
                }
            }
            sw.Stop();
            Console.WriteLine($"  NoLock (direct): {sw.ElapsedMilliseconds}ms");
            
            doc.Save();
        }
        
        // Expected results:
        // Default: ~500ms (lock overhead)
        // Sequential: ~250ms (no locks)
        // Batch: ~100ms (parallel preparation + single lock)
        // NoLock: ~200ms (direct access, no safety)
    }
}
```

### Object Insertion Examples

```csharp
internal static class ObjectInsertionExamples {
    public static void Example_InsertObjects_Basic(string folderPath, bool openExcel) {
        Console.WriteLine("[*] Excel - Insert objects (basic)");
        
        var people = new[] {
            new Person { Name = "Alice", Age = 30, Department = "IT" },
            new Person { Name = "Bob", Age = 25, Department = "HR" }
        };
        
        using (var document = ExcelDocument.Create(Path.Combine(folderPath, "Objects_Basic.xlsx"))) {
            var sheet = document.AddWorkSheet("People");
            
            // Basic insertion
            sheet.InsertObjects(people);
            
            document.Save(openExcel);
        }
    }
    
    public static void Example_InsertObjects_Advanced(string folderPath, bool openExcel) {
        Console.WriteLine("[*] Excel - Insert objects (advanced)");
        
        var people = GeneratePeople(5000);
        
        using (var document = ExcelDocument.Create(Path.Combine(folderPath, "Objects_Advanced.xlsx"))) {
            var sheet = document.AddWorkSheet("People");
            
            // Advanced insertion with options
            sheet.InsertObjects(people, new ObjectInsertionOptions {
                Mode = ExecutionMode.Parallel,  // Force parallel for 5000 objects
                ExpandNestedObjects = true,
                MaxNestingDepth = 2,
                ExcludeProperties = new List<string> { "InternalId" },
                HeaderTransform = header => header.Replace(".", " "),  // "Address.City" -> "Address City"
                CreateTable = true,
                TableStyle = TableStyle.Medium2
            });
            
            // No need to call AutoFit or AddTable - it's done automatically
            document.Save(openExcel);
        }
    }
    
    public static void Example_InsertObjects_NoLock(string folderPath, bool openExcel) {
        Console.WriteLine("[*] Excel - Insert objects (NoLock)");
        
        var data = GenerateLargeDictionaries(10000);
        
        using (var document = ExcelDocument.Create(Path.Combine(folderPath, "Objects_NoLock.xlsx"))) {
            var sheet = document.AddWorkSheet("Data");
            
            var sw = Stopwatch.StartNew();
            
            // Maximum performance with NoLock
            using (sheet.BeginNoLock()) {
                sheet.InsertObjectsNoLock(data, new ObjectInsertionOptions {
                    CreateTable = false  // Tables might need locks, skip for pure speed
                });
            }
            
            sw.Stop();
            Console.WriteLine($"  Inserted 10,000 objects in {sw.ElapsedMilliseconds}ms");
            
            document.Save(openExcel);
        }
    }
}
```

### Migration Examples (Show Before/After)

```csharp
internal static class MigrationExamples {
    /// <summary>
    /// BEFORE: Code that might deadlock or perform poorly
    /// </summary>
    public static void Example_Before() {
        using (var document = ExcelDocument.Create("before.xlsx")) {
            var sheet = document.AddWorkSheet("Data");
            
            // Problem 1: Individual cell operations with locks
            for (int i = 1; i <= 1000; i++) {
                sheet.CellValue(i, 1, $"Value {i}");  // 1000 lock operations!
            }
            
            // Problem 2: Multiple structural operations that might deadlock
            sheet.AutoFitColumns();
            sheet.AddAutoFilter("A1:A1000");
            sheet.AddConditionalColorScale("A1:A1000", Color.Red, Color.Green);
            
            document.Save();
        }
    }
    
    /// <summary>
    /// AFTER: Fixed code with better performance
    /// </summary>
    public static void Example_After() {
        using (var document = ExcelDocument.Create("after.xlsx")) {
            // Option 1: Quick fix - disable locks
            document.Execution.Mode = ExecutionMode.Sequential;
            
            var sheet = document.AddWorkSheet("Data");
            
            // Better: Use batch operation
            var cells = Enumerable.Range(1, 1000)
                .Select(i => (i, 1, (object)$"Value {i}"))
                .ToList();
            sheet.CellValues(cells);  // Single decision, optimized
            
            // Structural operations now work without deadlock
            sheet.AutoFitColumns();
            sheet.AddAutoFilter("A1:A1000");
            sheet.AddConditionalColorScale("A1:A1000", Color.Red, Color.Green);
            
            document.Save();
        }
    }
    
    /// <summary>
    /// ADVANCED: Power user optimization
    /// </summary>
    public static void Example_Advanced() {
        using (var document = ExcelDocument.Create("advanced.xlsx")) {
            var sheet = document.AddWorkSheet("Data");
            
            // Configure fine-grained control
            sheet.ExecutionOverride = new ExecutionPolicy {
                ParallelThreshold = 100,  // Lower threshold for this sheet
                OperationThresholds = {
                    ["CellValues"] = 50,      // Even lower for cell operations
                    ["AutoFitColumns"] = 5     // AutoFit benefits from parallel early
                }
            };
            
            // Use NoLock for maximum performance where safe
            using (sheet.BeginNoLock()) {
                // Bulk insert with no locks
                var cells = GenerateCells(10000);
                foreach (var (row, col, val) in cells) {
                    sheet.CellValueNoLock(row, col, val);
                }
                
                // Structural operations also lock-free
                sheet.AutoFitColumnsNoLock();
            }
            
            document.Save();
        }
    }
}
```

## Expected Outcomes

1. **Immediate fix**: Fluent API with multiple operations works without deadlocks
2. **Performance gain**: 50-100% faster for sequential operations (no lock overhead)
3. **User control**: Three levels of control (Auto/Sequential/Parallel)
4. **Backward compatible**: Existing code continues to work
5. **Power user path**: NoLock methods for maximum performance
6. **Clear semantics**: NoLock suffix explicitly indicates unsafe operations

## Notes

- The `AsyncLocal<bool>` pattern elegantly handles nested contexts without parameter passing
- OpenXML DOM is not thread-safe, so parallel operations must compute in parallel but apply sequentially
- The three-layer pattern (Safe/NoLock/Core) minimizes code duplication
- Operation-specific thresholds allow fine-tuning without code changes
- Diagnostic logging helps users understand and optimize their usage patterns