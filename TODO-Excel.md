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

## Proposed Solution

### Phase 1: Add User Control for Parallel Processing
Allow users to explicitly control whether operations use parallel processing:

#### 1.1 Document-Level Control
```csharp
public class ExcelDocument {
    /// <summary>
    /// Enables or disables parallel processing for all operations.
    /// Default: true
    /// </summary>
    public bool EnableParallelProcessing { get; set; } = true;
    
    /// <summary>
    /// Minimum number of cells to trigger parallel processing.
    /// Default: 1000
    /// </summary>
    public int ParallelThreshold { get; set; } = 1000;
}
```

#### 1.2 Sheet-Level Override
```csharp
public partial class ExcelSheet {
    /// <summary>
    /// Override parallel processing setting for this sheet.
    /// null = use document setting
    /// </summary>
    public bool? EnableParallelProcessing { get; set; }
}
```

#### 1.3 Method-Level Control
```csharp
public void CellValues(IEnumerable<(int Row, int Column, object Value)> cells, bool? useParallel = null)
public void InsertDataTable(DataTable dataTable, int startRow, int startColumn, bool? useParallel = null)
```

### Phase 2: Separate Internal Implementation (No Locks)
Create internal methods that don't use locks for sequential operations:

#### 2.1 Option A: Expose Unsafe/Direct Methods to Users
```csharp
// Give power users direct access to non-locked methods
// They take responsibility for thread safety

public partial class ExcelSheet {
    // Safe API (current, with locks)
    public void CellValue(int row, int column, string value) {
        WriteLock(() => CellValueDirect(row, column, value));
    }
    
    // Direct API (new, no locks) - clearly marked as unsafe
    public void CellValueDirect(int row, int column, string value) {
        Cell cell = GetCell(row, column);
        int sharedStringIndex = _excelDocument.GetSharedStringIndex(value);
        cell.CellValue = new CellValue(sharedStringIndex.ToString());
        cell.DataType = CellValues.SharedString;
    }
    
    // Or use a more explicit naming convention:
    public void CellValueUnsafe(int row, int column, string value) { ... }
    public void CellValueNoLock(int row, int column, string value) { ... }
}

// Usage by power users who know what they're doing:
// They ensure thread safety at a higher level
using (sheet.BeginExclusiveAccess()) {  // Optional helper for explicit locking
    for (int i = 0; i < 10000; i++) {
        sheet.CellValueDirect(i, 1, $"Value {i}");  // No lock overhead!
    }
}
```

#### 2.2 Option B: Unsafe Context Pattern
```csharp
// Provide an "unsafe" context where all operations skip locks
public class UnsafeSheetContext {
    private readonly ExcelSheet _sheet;
    
    internal UnsafeSheetContext(ExcelSheet sheet) {
        _sheet = sheet;
    }
    
    // All methods here are direct/no-lock versions
    public void CellValue(int row, int column, object value) {
        // Direct implementation, no locks
    }
    
    public void AutoFitColumns() {
        // Direct implementation, no locks
    }
}

// Usage:
var unsafe = sheet.Unsafe();  // or sheet.Direct() or sheet.NoLock()
for (int i = 0; i < 10000; i++) {
    unsafe.CellValue(i, 1, $"Value {i}");  // No locks!
}
```

#### 2.3 Option C: Batch-Level Decision (Original Plan)
```csharp
// WRONG APPROACH (too many checks):
// DON'T check locking on every cell - that's thousands of if statements!

// RIGHT APPROACH - Decide once at batch level:
public void CellValues(IEnumerable<(int Row, int Column, object Value)> cells, bool? useParallel = null) {
    var cellList = cells.ToList();
    bool shouldUseParallel = useParallel ?? (EnableParallelProcessing && cellList.Count > ParallelThreshold);
    
    if (shouldUseParallel) {
        // Parallel path with ONE lock acquisition for entire batch
        CellValuesParallelImpl(cellList);
    } else {
        // Sequential path with NO locks at all
        CellValuesSequentialNoLock(cellList);
    }
}

// Sequential implementation - NEVER uses locks
private void CellValuesSequentialNoLock(List<(int Row, int Column, object Value)> cells) {
    // Direct DOM manipulation, no locks needed
    foreach (var (row, column, value) in cells) {
        Cell cell = GetCell(row, column);
        // ... set value directly
    }
}

// Keep individual CellValue for backward compatibility
public void CellValue(int row, int column, string value) {
    // Single cells always use locks for thread safety
    // But users should batch operations when possible
    WriteLock(() => SetCellValueInternal(row, column, value));
}
```

#### 2.4 Pros and Cons of Each Approach

**Option A: Direct/Unsafe Methods**
- ✅ Maximum performance and control for users
- ✅ Simple implementation (minimal redundancy - just thin wrappers)
- ✅ Clear separation between safe and unsafe
- ❌ Users can easily shoot themselves in the foot
- ❌ More API surface area to maintain

**Option B: Unsafe Context**
- ✅ Clean API separation
- ✅ Harder to accidentally use unsafe methods
- ✅ Can batch multiple operations in unsafe context
- ❌ Some code duplication (need to mirror methods)
- ❌ Another class to maintain

**Option C: Batch-Level Decision**
- ✅ Safest - users can't break thread safety
- ✅ Automatic optimization
- ❌ Less control for power users
- ❌ Can't optimize small operations as well

#### 2.5 Recommended Approach: Combine A + C
```csharp
public partial class ExcelSheet {
    // 1. Keep safe API as-is (backward compatible)
    public void CellValue(int row, int column, object value) {
        WriteLock(() => CellValueDirect(row, column, value));
    }
    
    // 2. Expose direct methods for power users (new)
    public void CellValueDirect(int row, int column, object value) {
        // Implementation without locks
        // Minimal code duplication - this IS the implementation
    }
    
    // 3. Smart batch methods (auto-optimize)
    public void CellValues(IEnumerable<(int, int, object)> cells, bool? useParallel = null) {
        var cellList = cells.ToList();
        bool shouldUseParallel = useParallel ?? (EnableParallelProcessing && cellList.Count > ParallelThreshold);
        
        if (shouldUseParallel) {
            CellValuesParallelImpl(cellList);  // With locks
        } else {
            // Just call the direct methods!
            foreach (var (row, col, val) in cellList) {
                CellValueDirect(row, col, val);  // No locks
            }
        }
    }
}
```

This gives users three levels of control:
1. **Safe default**: `CellValue()` - always locked
2. **Power user**: `CellValueDirect()` - no locks, full control
3. **Smart batching**: `CellValues()` - auto-optimized

### Phase 3: Fix Nested Locking Issues
Ensure consistent handling of recursive locks:

#### 3.1 Standardize Lock Checking
- All methods that might be called within a lock should check `_lock.IsWriteLockHeld`
- Create helper method: `ExecuteWithOptionalLock(Action action)`

#### 3.2 Problem Methods to Fix
- [ ] `AutoFitColumns()` - calls `AutoFitColumn()` in a loop
- [ ] `AutoFitRows()` - calls `AutoFitRow()` in a loop  
- [ ] `AddAutoFilter()` - modifies worksheet structure
- [ ] `AddConditionalColorScale()` - modifies worksheet structure
- [ ] `AddConditionalDataBar()` - modifies worksheet structure
- [ ] `SetColumnWidth()` / `SetColumnHidden()` - column modifications

### Phase 4: Optimize Batch Operations
Reduce lock acquisitions for batch operations:

#### 4.1 Encourage Batching Over Individual Calls
```csharp
// SLOW (1000 lock acquisitions):
for (int i = 0; i < 1000; i++) {
    sheet.CellValue(i, 1, $"Value {i}");  // Each call acquires/releases lock
}

// FAST (0 locks for sequential, 1 lock for parallel):
var cells = Enumerable.Range(1, 1000).Select(i => (i, 1, (object)$"Value {i}"));
sheet.CellValues(cells, useParallel: false);  // No locks at all!

// Or let it auto-decide:
sheet.CellValues(cells);  // Will choose based on count and settings
```

#### 4.2 Fluent API Optimization
```csharp
// The fluent API should internally batch operations:
public SheetBuilder Row(Action<RowBuilder> action) {
    var builder = new RowBuilder(this, Sheet, _currentRow);
    action(builder);
    
    // Instead of individual CellValue calls, batch them:
    if (_batchMode) {
        _pendingCells.AddRange(builder.GetCells());
    } else {
        builder.ApplyValues();  // Current behavior
    }
    return this;
}

// Then apply all at once when leaving fluent context:
public ExcelDocument End() {
    if (_pendingCells.Any()) {
        Sheet.CellValues(_pendingCells, useParallel: false);  // No locks!
    }
    return Workbook;
}
```

## Phase 5: Enhanced Object/Data Insertion
Improve the existing `InsertObjects` functionality with more control:

### 5.1 Enhanced Object Insertion API
```csharp
public class ObjectInsertionOptions {
    public bool IncludeHeaders { get; set; } = true;
    public int StartRow { get; set; } = 1;
    public int StartColumn { get; set; } = 1;
    public bool? UseParallel { get; set; }  // null = auto-decide
    public bool ExpandNestedObjects { get; set; } = true;
    public int MaxNestingDepth { get; set; } = 3;
    public List<string> IncludeProperties { get; set; }  // Whitelist
    public List<string> ExcludeProperties { get; set; }  // Blacklist
    public Func<string, string> HeaderTransform { get; set; }  // Custom header names
    public bool CreateTable { get; set; } = false;
    public TableStyle TableStyle { get; set; } = TableStyle.Medium2;
}

// Usage examples:
// 1. Simple objects
var people = new[] {
    new { Name = "Alice", Age = 30, Address = new { City = "NYC", Zip = "10001" } },
    new { Name = "Bob", Age = 25, Address = new { City = "LA", Zip = "90001" } }
};

sheet.InsertObjects(people, new ObjectInsertionOptions {
    ExpandNestedObjects = true,  // Creates columns: Name, Age, Address.City, Address.Zip
    UseParallel = false,  // Force sequential for small datasets
    CreateTable = true    // Automatically create Excel table
});

// 2. Hashtables/Dictionaries
var data = new[] {
    new Dictionary<string, object> { ["ID"] = 1, ["Status"] = "Active" },
    new Dictionary<string, object> { ["ID"] = 2, ["Status"] = "Pending" }
};
sheet.InsertObjects(data);

// 3. DataTable (already supported)
sheet.InsertDataTable(dataTable, 1, 1, useParallel: false);
```

### 5.2 Direct/NoLock Version for Object Insertion
```csharp
public void InsertObjectsDirect<T>(IEnumerable<T> items, ObjectInsertionOptions options = null) {
    // Process objects and flatten
    var cells = PrepareObjectCells(items, options);
    
    // Insert without any locks
    foreach (var (row, col, value) in cells) {
        CellValueDirect(row, col, value);  // No locks!
    }
    
    // Optional: Create table without locks
    if (options?.CreateTable == true) {
        AddTableDirect(range, hasHeaders: true, name, style);
    }
}
```

## Phase 6: Fix All Methods with Similar Locking Issues

### 6.1 Methods That Need Direct/NoLock Versions
All these methods currently use `WriteLock` and need direct versions:

```csharp
// Column/Row operations
public void SetColumnWidthDirect(int columnIndex, double width)
public void SetColumnHiddenDirect(int columnIndex, bool hidden)
public void SetRowHeightDirect(int rowIndex, double height)
public void AutoFitColumnsDirect()
public void AutoFitRowsDirect()

// Formatting operations  
public void FreezeDirect(int topRows = 0, int leftCols = 0)
public void AddAutoFilterDirect(string range, Dictionary<uint, IEnumerable<string>> criteria = null)
public void AddConditionalColorScaleDirect(string range, Color startColor, Color endColor)
public void AddConditionalDataBarDirect(string range, Color color)

// Table operations
public void AddTableDirect(string range, bool hasHeader, string name, TableStyle style)
public void FormatCellDirect(int row, int column, string numberFormat)
public void CellFormulaDirect(int row, int column, string formula)
```

### 6.2 Pattern for All Methods
```csharp
// Safe version (existing)
public void MethodName(params) {
    WriteLock(() => MethodNameDirect(params));
}

// Direct version (new)
public void MethodNameDirect(params) {
    // Actual implementation without locks
}
```

## Implementation Tasks

### High Priority (Fixes Current Issue)
1. [ ] Add `EnableParallelProcessing` property to `ExcelDocument`
2. [ ] Add `EnableParallelProcessing` property to `ExcelSheet`
3. [ ] Modify `CellValuesParallel` to check these settings
4. [ ] Update fluent API to respect parallel settings
5. [ ] Fix `AutoFitColumns/Rows` nested locking
6. [ ] Test with problematic fluent example

### Medium Priority (Performance)
7. [ ] Create internal methods without locks
8. [ ] Add `useParallel` parameter to batch methods
9. [ ] Implement `BeginBatch/EndBatch` API
10. [ ] Add `ShouldUseLocking()` helper method

### Low Priority (Polish)
11. [ ] Add configuration for parallel threshold per operation type
12. [ ] Add diagnostics/logging for lock contentions
13. [ ] Document best practices for parallel vs sequential
14. [ ] Add unit tests for all combinations

## Testing Plan

### Test Cases
1. **Original failing case**: Fluent API with AutoFilter + ConditionalFormatting + AutoFit
2. **Sequential mode**: Same operations with `EnableParallelProcessing = false`
3. **Large dataset**: > 10,000 cells with parallel enabled/disabled
4. **Concurrent access**: Multiple threads writing to different sheets
5. **Nested operations**: AutoFit inside batch operations

### Performance Benchmarks
- Measure lock overhead: Sequential with locks vs without locks
- Compare parallel vs sequential for various data sizes
- Profile memory usage in both modes

## Migration Guide for Users

### Quick Fix for Current Issues
```csharp
// If experiencing issues with Excel not opening:
using (var document = ExcelDocument.Create(filePath)) {
    document.EnableParallelProcessing = false;  // Disable parallel processing
    
    // Your existing code works without changes
    document.AsFluent()
        .Sheet("Data", s => s
            .AutoFilter("A1:B3")
            .ConditionalColorScale("B2:B3", Color.Red, Color.Green)
            .AutoFit(columns: true, rows: true)
        )
        .End()
        .Save(openExcel);
}
```

### Performance Optimization
```csharp
// For large datasets, explicitly enable parallel:
sheet.CellValues(largeCellCollection, useParallel: true);

// For small datasets or when debugging:
sheet.CellValues(smallCellCollection, useParallel: false);
```

## Notes
- The current `ReaderWriterLockSlim` with `LockRecursionPolicy.SupportsRecursion` has overhead
- Consider replacing with simpler locking mechanism for non-recursive scenarios
- Parallel processing benefits are minimal for < 1000 cells due to overhead
- Most fluent API usage involves small datasets that don't benefit from parallelization